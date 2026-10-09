(function () {
    'use strict';
    var root = document.querySelector('.nsr-admin');
    if (!root) return;
    // Preserve the existing table and event handlers; only change its presentation.
    root.querySelectorAll('table.widefat').forEach(function (table) {
        var headers = Array.from(table.querySelectorAll('thead th')).map(function (th) { return th.textContent.trim(); });
        if (!headers.length) return;
        table.classList.add('nsr-mobile-cards');
        table.querySelectorAll('tbody tr').forEach(function (row) {
            Array.from(row.children).forEach(function (cell, index) {
                cell.dataset.label = headers[index] || '';
            });
        });
    });
    var media = window.matchMedia('(max-width: 782px)');
    function adapt() {
        root.querySelectorAll('[data-mobile-collapse]').forEach(function (details) { details.open = !media.matches; });
    }
    adapt();
    media.addEventListener('change', adapt);
})();
