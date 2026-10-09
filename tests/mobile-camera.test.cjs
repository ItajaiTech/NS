const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
function setup(options = {}) {
    const elements = {};
    const el = name => elements[name] ||= {value:'', disabled:false, hidden:true, textContent:'', handlers:{}, addEventListener(k,fn){this.handlers[k]=fn;}};
    const root = {isConnected:true, querySelector:el};
    const video = el('video');
    Object.assign(video,{readyState:2,videoWidth:640,videoHeight:480,play:async()=>{}});
    const tracks = {stopped:0, stop(){this.stopped++;}};
    const stream = {getTracks:()=>[tracks]};
    const sent = [], events = {};
    const sku = {value:'SKU1'};
    const window = {isSecureContext:options.secure !== false, nsrScanMultipleNs:ns=>sent.push(ns), addEventListener:(key,fn)=>events[key]=fn,
        jsQR:()=>({data:options.qr || '00123\n00124'})};
    const document = {hidden:false,body:{},getElementById:id=>id==='nsr-camera'?root:sku,querySelectorAll:()=>[],addEventListener:(key,fn)=>events[key]=fn,
        createElement:()=>({getContext:()=>({drawImage(){},getImageData:()=>({data:[]})})})};
    vm.runInNewContext(fs.readFileSync('ns-rastreio-plugin/assets/mobile-camera.js','utf8'),{window,document,navigator:{mediaDevices:{getUserMedia:options.open || (async()=>stream)}},
        MutationObserver:class {observe(){}},setTimeout,clearTimeout});
    return {el,tracks,sent,events,sku};
}
test('QR pauses camera, preserves zeros and registers only after confirmation', async()=>{
    const h=setup();
    await h.el('[data-start]').handlers.click();
    assert.equal(h.tracks.stopped,1);
    assert.equal(h.sent.length,0);
    assert.equal(h.el('textarea').value,'00123\n00124');
    h.el('[data-confirm]').handlers.click();
    assert.deepEqual(Array.from(h.sent[0]),['00123','00124']);
    await h.el('[data-start]').handlers.click();
    h.el('[data-confirm]').handlers.click();
    assert.equal(h.sent.length,1);
});
test('insecure site and denied permission show actionable messages', async()=>{
    const insecure=setup({secure:false});
    await insecure.el('[data-start]').handlers.click();
    assert.match(insecure.el('[role="status"]').textContent,/HTTPS/);
    const denied=setup({open:async()=>{throw Object.assign(new Error(),{name:'NotAllowedError'});}});
    await denied.el('[data-start]').handlers.click();
    assert.match(denied.el('[role="status"]').textContent,/Permita/);
    assert.equal(denied.el('[data-start]').disabled,false);
});
test('camera opened after cancellation is released; QR links never reach queue',async()=>{
    let resolve;
    const h=setup({open:()=>new Promise(r=>resolve=r)});
    const pending=h.el('[data-start]').handlers.click();
    h.el('[data-stop]').handlers.click();
    let stopped=0;
    resolve({getTracks:()=>[{stop(){stopped++;}}]});
    await pending;
    assert.equal(stopped,1);
    const link=setup({qr:'https://example.com/123'});
    await link.el('[data-start]').handlers.click();
    link.el('[data-confirm]').handlers.click();
    assert.equal(link.sent.length,0);
});
