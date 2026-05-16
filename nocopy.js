(function () {
    document.addEventListener('contextmenu', function (e) { e.preventDefault(); });
    document.addEventListener('copy',        function (e) { e.preventDefault(); });
    document.addEventListener('cut',         function (e) { e.preventDefault(); });
    document.addEventListener('selectstart', function (e) { e.preventDefault(); });
    document.addEventListener('dragstart',   function (e) { e.preventDefault(); });
})();

(function () {
    var triggered = false;

    function sizeCheck() {
        return window.outerWidth  - window.innerWidth  > 160 ||
               window.outerHeight - window.innerHeight > 160;
    }

    function debugCheck() {
        var t = performance.now();
        (function () { debugger; }());
        return performance.now() - t > 100;
    }

    function block() {
        if (triggered) return;
        triggered = true;
        document.documentElement.innerHTML =
            '<body style="margin:0;display:flex;align-items:center;justify-content:center;' +
            'height:100vh;background:#0d0b1e;font-family:sans-serif;">' +
            '<p style="color:#f87171;font-size:18px;font-weight:700;">&#9888; Akses tidak diizinkan</p>' +
            '</body>';
        setTimeout(function () { window.location.replace('/'); }, 1500);
    }

    function check() {
        if (sizeCheck() || debugCheck()) block();
    }

    setInterval(check, 1000);
    window.addEventListener('resize', check);
}());
