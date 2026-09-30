const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('ns-rastreio-plugin/ns-rastreio-plugin.php', 'utf8');
function setup() {
    const elements = {};
    const get = id => elements[id] ||= {value:'', focus(){}, addEventListener(){}};
    const requests = [], feedback = [], updates = [], timers = [];
    const ctx = vm.createContext({window:{}, document:{getElementById:get}, FormData,
        activeSku:'PRD00017', TOKEN:'token', NONCE:'nonce', AJAXURL:'/ajax',
        showFeedback:(...args)=>feedback.push(args), updateRow:(...args)=>updates.push(args),
        setTimeout:fn=>(timers.push(fn), timers.length), clearTimeout:id=>{if(id) timers[id-1]=()=>{};},
        fetch:(url, options)=>new Promise((resolve,reject)=>requests.push({body:options.body, resolve, reject}))
    });
    let start = source.indexOf('                // Capture immediately;');
    let end = source.indexOf('                // ----- Gerar sequencia', start);
    vm.runInContext(source.slice(start,end),ctx);
    start=source.indexOf('                window.nsrScanMultipleNs =');
    end=source.indexOf('                // Event listener: paste',start);
    vm.runInContext(source.slice(start,end),ctx);
    const enter = ns=>{get('nsr-inp-ns').value=ns; ctx.window.nsrScanNs();};
    const respond = async(index, result)=>{
        requests[index].resolve({json:async()=>result || {success:true,data:{sku:'PRD00017',ns:requests[index].body.get('ns'),scanned_count:index+1,expected:10,is_ok:index===9}}});
        await new Promise(setImmediate);
    };
    return {ctx,get,requests,feedback,updates,timers,enter,respond};
}
test('QR with ten lines and CRLF: exact serials, one request at a time, preserves next input',async()=>{
    const h=setup();
    const serials=['3222909','3222914','3222910','3222915','3222911','3222907','3222903','3222908','3222904','3222913'];
    for(const ns of serials){h.enter(ns); assert.equal(h.get('nsr-inp-ns').value,''); h.enter('');}
    assert.equal(h.requests.length,1);
    assert.equal(vm.runInContext('scansPending()',h.ctx),true);
    h.get('nsr-inp-ns').value='NEXT-PARTIAL';
    for(let i=0;i<serials.length;i++){
        assert.equal(h.requests.length,i+1);
        assert.equal(h.requests[i].body.get('ns'),serials[i]);
        assert.equal(h.requests[i].body.get('sku'),'PRD00017');
        await h.respond(i);
        assert.equal(h.get('nsr-inp-ns').value,'NEXT-PARTIAL');
    }
    assert.equal(vm.runInContext('scansPending()',h.ctx),false);
    h.timers.forEach(fn=>fn());
    assert.equal(h.updates.length,10);
    assert.ok(h.updates.every(update=>update[1]===false));
});
test('pasted batches share queue and retain captured SKU and leading zeroes',async()=>{
    const h=setup();
    h.ctx.window.nsrScanMultipleNs(['00123','00124']);
    h.ctx.activeSku='SECOND';
    h.ctx.window.nsrScanMultipleNs(['ABC-42']);
    await h.respond(0); await h.respond(1); await h.respond(2);
    assert.deepEqual(h.requests.map(r=>[r.body.get('ns'),r.body.get('sku')]),[['00123','PRD00017'],['00124','PRD00017'],['ABC-42','SECOND']]);
});
test('server and network failures retain serial details and continue queue without clearing typing',async()=>{
    const h=setup();
    h.ctx.window.nsrScanMultipleNs(['DUP','NETWORK','GOOD']);
    await h.respond(0,{success:false,data:{msg:'Duplicado'}});
    h.requests[1].reject(new Error('offline')); await new Promise(setImmediate);
    await h.respond(2);
    const last=h.feedback.at(-1);
    assert.equal(last[1],'error'); assert.match(last[0],/DUP: Duplicado/); assert.match(last[0],/NETWORK: falha/);
    assert.equal(h.requests.length,3);
});
test('rendered inline scan script has valid JavaScript syntax',()=>{
    const start=source.lastIndexOf('<script>',source.indexOf('var activeSku = null;'))+8;
    const end=source.indexOf('</script>',start);
    new vm.Script(source.slice(start,end).replace(/<\?php[\s\S]*?\?>/g,'"test"'));
});
