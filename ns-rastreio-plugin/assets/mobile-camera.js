(function () {
    'use strict';
    var root = document.getElementById('nsr-camera');
    if (!root) return;
    var video = root.querySelector('video');
    var status = root.querySelector('[role="status"]');
    var preview = root.querySelector('textarea');
    var start = root.querySelector('[data-start]');
    var confirm = root.querySelector('[data-confirm]');
    var stream = null, timer = null, generation = 0, lastAccepted = '';
    var canvas = document.createElement('canvas');
    var context = canvas.getContext('2d', {willReadFrequently: true});

    function stop() {
        generation++;
        clearTimeout(timer);
        if (stream) stream.getTracks().forEach(function (track) { track.stop(); });
        stream = null;
        video.srcObject = null;
        video.hidden = true;
        start.disabled = false;
    }
    async function read(id) {
        if (id !== generation || !stream) return;
        try {
            if (video.readyState >= 2 && video.videoWidth) {
                var scale = Math.min(1, 800 / video.videoWidth);
                canvas.width = Math.round(video.videoWidth * scale);
                canvas.height = Math.round(video.videoHeight * scale);
                context.drawImage(video, 0, 0, canvas.width, canvas.height);
                var frame = context.getImageData(0, 0, canvas.width, canvas.height);
                var code = window.jsQR(frame.data, canvas.width, canvas.height);
                if (code && code.data.trim()) {
                    stop();
                    preview.value = code.data.trim();
                    confirm.disabled = false;
                    status.textContent = 'QR lido. Confira os NS e o SKU antes de registrar.';
                    if (navigator.vibrate) navigator.vibrate(80);
                    return;
                }
            }
            timer = setTimeout(function () { read(id); }, 180);
        } catch (error) {
            stop();
            status.textContent = 'Falha ao ler a camera. Tente novamente ou digite o NS.';
        }
    }
    start.addEventListener('click', async function () {
        stop();
        if (!window.isSecureContext) {
            status.textContent = 'Abra o site em HTTPS para utilizar a camera.';
            return;
        }
        if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia || !window.jsQR) {
            status.textContent = 'Leitor indisponivel. Atualize o navegador ou informe o NS manualmente.';
            return;
        }
        if (!document.getElementById('nsr-inp-sku').value) {
            status.textContent = 'Selecione um SKU antes de abrir a camera.';
            return;
        }
        preview.value = '';
        confirm.disabled = true;
        start.disabled = true;
        status.textContent = 'Abrindo camera traseira...';
        var id = generation;
        try {
            var opened = await navigator.mediaDevices.getUserMedia({audio: false, video: {facingMode: {ideal: 'environment'}}});
            if (id !== generation) {
                opened.getTracks().forEach(function (track) { track.stop(); });
                return;
            }
            stream = opened;
            video.srcObject = stream;
            video.hidden = false;
            await video.play();
            if (id !== generation) return;
            status.textContent = 'Aponte a camera para o QR code dos numeros de serie.';
            read(id);
        } catch (error) {
            if (id !== generation) return;
            stop();
            status.textContent = error.name === 'NotAllowedError'
                ? 'Permita o acesso a camera nas configuracoes do navegador e tente novamente.'
                : 'Nao foi possivel abrir a camera. Confira se ela esta disponivel e tente novamente.';
        }
    });
    root.querySelector('[data-stop]').addEventListener('click', function () {
        stop();
        status.textContent = 'Camera fechada.';
    });
    confirm.addEventListener('click', function () {
        var text = preview.value.trim();
        if (!text || typeof window.nsrScanMultipleNs !== 'function') return;
        if (/https?:\/\/|[{}<>]/i.test(text)) {
            status.textContent = 'Este QR contem um link ou dados estruturados. Informe somente os numeros de serie no campo acima.';
            return;
        }
        var sku = document.getElementById('nsr-inp-sku').value;
        if (!sku) { status.textContent = 'Selecione um SKU.'; return; }
        var key = sku + '\n' + text;
        if (key === lastAccepted) {
            status.textContent = 'Este QR ja foi enviado para este SKU. Confira a lista de NS e o resultado da gravacao.';
            return;
        }
        lastAccepted = key;
        window.nsrScanMultipleNs(text.split(/[\r\n,;|\s]+/));
        confirm.disabled = true;
        status.textContent = 'NS enviados para gravacao. Confira o resultado da bipagem abaixo antes de ler o proximo QR.';
    });
    document.addEventListener('visibilitychange', function () { if (document.hidden) stop(); });
    window.addEventListener('pagehide', stop);
    document.querySelectorAll('.nsr-navigation a').forEach(function (tab) { tab.addEventListener('click', stop); });
    new MutationObserver(function () { if (!root.isConnected) stop(); }).observe(document.body, {childList: true, subtree: true});
})();
