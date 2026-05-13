(function () {
    'use strict';

    /* ─── Inject CSS ─── */
    const css = `
        .otp-overlay {
            position: fixed; inset: 0; z-index: 9999;
            background: rgba(8,6,30,0.88);
            backdrop-filter: blur(8px);
            display: flex; align-items: center; justify-content: center;
            padding: 20px; box-sizing: border-box;
        }
        .otp-card {
            background: linear-gradient(145deg, #1a1043, #120d35);
            border: 1px solid rgba(124,58,237,0.45);
            border-radius: 22px;
            padding: 40px 32px 32px;
            width: 100%; max-width: 380px;
            box-shadow: 0 24px 64px rgba(0,0,0,0.6), 0 0 0 1px rgba(124,58,237,0.1);
            text-align: center;
        }
        .otp-icon { font-size: 38px; margin-bottom: 14px; }
        .otp-title {
            font-size: 22px; font-weight: 800; color: white;
            margin: 0 0 8px; letter-spacing: -0.3px;
        }
        .otp-sub {
            font-size: 13px; color: #a78bfa;
            line-height: 1.6; margin-bottom: 22px;
        }
        .otp-sub strong { color: white; }
        .otp-dev-box {
            background: rgba(251,191,36,0.1);
            border: 1px solid rgba(251,191,36,0.35);
            border-radius: 10px; padding: 10px 16px;
            margin-bottom: 18px; font-size: 14px;
            color: #fbbf24; font-weight: 700;
            letter-spacing: 1px;
        }
        .otp-dev-label {
            font-size: 10px; color: #d97706;
            text-transform: uppercase; letter-spacing: 1px;
            display: block; margin-bottom: 4px;
        }
        .otp-input {
            width: 100%; padding: 14px 10px;
            background: rgba(255,255,255,0.07);
            border: 2px solid rgba(124,58,237,0.4);
            border-radius: 12px; color: white;
            font-size: 28px; font-weight: 800;
            letter-spacing: 12px; text-align: center;
            outline: none; box-sizing: border-box;
            margin-bottom: 10px; transition: 0.2s;
            font-variant-numeric: tabular-nums;
        }
        .otp-input:focus { border-color: #7c3aed; background: rgba(124,58,237,0.1); }
        .otp-error {
            font-size: 12px; color: #f87171;
            margin-bottom: 12px; min-height: 18px;
            line-height: 1.4;
        }
        .otp-verify-btn {
            width: 100%; padding: 14px;
            background: #7c3aed; color: white;
            border: none; border-radius: 12px;
            font-size: 15px; font-weight: 700;
            cursor: pointer; letter-spacing: 2px;
            transition: 0.2s; margin-bottom: 18px;
        }
        .otp-verify-btn:hover:not(:disabled) {
            background: #a855f7;
            box-shadow: 0 6px 20px rgba(124,58,237,0.45);
        }
        .otp-verify-btn:disabled { opacity: 0.6; cursor: not-allowed; }
        .otp-resend-area { font-size: 12px; color: #6b7280; }
        .otp-resend-btn {
            background: none; border: none;
            color: #a78bfa; font-size: 12px;
            font-weight: 700; cursor: pointer;
            text-decoration: underline; padding: 0;
        }
        .otp-resend-btn:hover { color: white; }
        .otp-check-loading {
            display: flex; align-items: center; justify-content: center;
            gap: 8px; color: #a78bfa; font-size: 13px;
            padding: 20px 0;
        }
    `;
    const styleEl = document.createElement('style');
    styleEl.textContent = css;
    document.head.appendChild(styleEl);

    /* ─── Cookie helpers ─── */
    function setCookie(name, val, days) {
        const exp = new Date(Date.now() + days * 864e5).toUTCString();
        document.cookie = name + '=' + encodeURIComponent(val) +
            '; expires=' + exp + '; path=/; SameSite=Strict';
    }

    function getCookie(name) {
        const m = document.cookie.match('(?:^|;)\\s*' + name + '=([^;]*)');
        return m ? decodeURIComponent(m[1]) : null;
    }

    /* ─── API helpers ─── */
    function post(url, body) {
        return fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        }).then(function (r) { return r.json(); });
    }

    /* ─── Device check ─── */
    function checkDevice(phone) {
        var token  = getCookie('adk_dtoken');
        var dphone = getCookie('adk_dphone');
        if (!token || dphone !== phone) return Promise.resolve(false);
        return post('api/check_device.php', { phone: phone, token: token })
            .then(function (d) { return !!d.ok; })
            .catch(function () { return false; });
    }

    /* ─── Modal state ─── */
    var overlay      = null;
    var cdInterval   = null;

    function buildOverlay() {
        overlay = document.createElement('div');
        overlay.className = 'otp-overlay';
        overlay.innerHTML =
            '<div class="otp-card">' +
                '<div class="otp-icon">🔒</div>' +
                '<div class="otp-title">Verifikasi OTP</div>' +
                '<div class="otp-sub">Kode dikirim ke<br><strong id="otp-phone-disp"></strong></div>' +
                '<div id="otp-dev-box" class="otp-dev-box" style="display:none">' +
                    '<span class="otp-dev-label">DEV MODE — Kode OTP</span>' +
                    '<span id="otp-dev-code"></span>' +
                '</div>' +
                '<input type="tel" id="otp-input" class="otp-input" maxlength="6"' +
                '       placeholder="000000" inputmode="numeric" autocomplete="one-time-code">' +
                '<div id="otp-error" class="otp-error"></div>' +
                '<button id="otp-btn" class="otp-verify-btn">VERIFIKASI</button>' +
                '<div class="otp-resend-area">' +
                    '<span id="otp-cd-wrap">Kirim ulang dalam <strong id="otp-cd">60</strong>s</span>' +
                    '<button id="otp-resend" class="otp-resend-btn" style="display:none">Kirim Ulang</button>' +
                '</div>' +
            '</div>';
        document.body.appendChild(overlay);
    }

    function showDev(code) {
        var box  = document.getElementById('otp-dev-box');
        var span = document.getElementById('otp-dev-code');
        if (box && span) { span.textContent = code; box.style.display = 'block'; }
    }

    function setError(msg) {
        var el = document.getElementById('otp-error');
        if (el) el.textContent = msg || '';
    }

    function startCD(phone, onResend) {
        clearInterval(cdInterval);
        var left   = 60;
        var cdEl   = document.getElementById('otp-cd');
        var cdWrap = document.getElementById('otp-cd-wrap');
        var resBtn = document.getElementById('otp-resend');

        if (cdWrap) cdWrap.style.display = 'inline';
        if (resBtn) resBtn.style.display = 'none';

        cdInterval = setInterval(function () {
            left--;
            if (cdEl) cdEl.textContent = left;
            if (left <= 0) {
                clearInterval(cdInterval);
                if (cdWrap) cdWrap.style.display = 'none';
                if (resBtn) resBtn.style.display  = 'inline';
            }
        }, 1000);

        if (resBtn) {
            resBtn.onclick = function () { onResend(); };
        }
    }

    function openModal(phone, onSuccess) {
        if (!overlay) buildOverlay();

        /* reset tampilan */
        document.getElementById('otp-phone-disp').textContent = phone;
        document.getElementById('otp-input').value  = '';
        document.getElementById('otp-error').textContent = '';
        document.getElementById('otp-dev-box').style.display = 'none';
        overlay.style.display = 'flex';

        /* set mode input — password tersembunyi jika email, numerik jika nomor HP */
        var inp       = document.getElementById('otp-input');
        var isEmail   = phone.indexOf('@') !== -1;

        if (isEmail) {
            inp.type        = 'password';
            inp.maxLength   = 100;
            inp.placeholder = '••••••••';
            inp.inputMode   = 'text';
            inp.style.letterSpacing = '2px';
            inp.style.fontSize      = '18px';
            inp.style.textAlign     = 'center';
            inp.oninput = function () { setError(''); };
        } else {
            inp.type        = 'tel';
            inp.maxLength   = 6;
            inp.placeholder = '000000';
            inp.inputMode   = 'numeric';
            inp.style.letterSpacing = '12px';
            inp.style.fontSize      = '28px';
            inp.style.textAlign     = 'center';
            inp.oninput = function () {
                this.value = this.value.replace(/\D/g, '').slice(0, 6);
                setError('');
            };
        }

        function doSend() {
            post('api/send_otp.php', { phone: phone }).then(function (d) {
                if (d.otp)   showDev(d.otp);
                if (d.error) setError(d.error);
            });
        }

        doSend();
        startCD(phone, function () { doSend(); startCD(phone, arguments.callee); });

        /* tombol verifikasi — hapus listener lama agar tidak double */
        var btn = document.getElementById('otp-btn');
        var newBtn = btn.cloneNode(true);
        btn.parentNode.replaceChild(newBtn, btn);

        function doVerify() {
            var code = document.getElementById('otp-input').value.trim();
            if (code.length < 6) { setError('Masukkan 6 digit kode OTP.'); return; }

            newBtn.disabled    = true;
            newBtn.textContent = 'Memverifikasi...';

            post('api/verify_otp.php', { phone: phone, code: code })
                .then(function (d) {
                    if (d.ok) {
                        if (d.redirect) {
                            window.location.href = d.redirect;
                            return;
                        }
                        setCookie('adk_dtoken', d.token, 30);
                        setCookie('adk_dphone', phone,   30);
                        overlay.style.display = 'none';
                        clearInterval(cdInterval);
                        onSuccess();
                    } else if (d.locked && d.remaining) {
                        startLockoutCountdown(d.remaining, newBtn, document.getElementById('otp-input'));
                    } else {
                        setError(d.error || 'Kode salah atau sudah kedaluwarsa.');
                        newBtn.disabled    = false;
                        newBtn.textContent = 'VERIFIKASI';
                    }
                })
                .catch(function () {
                    setError('Gagal terhubung ke server.');
                    newBtn.disabled    = false;
                    newBtn.textContent = 'VERIFIKASI';
                });
        }

        newBtn.onclick = doVerify;
        inp.onkeydown  = function (e) { if (e.key === 'Enter') doVerify(); };
        inp.focus();
    }

    /* ─── Lockout countdown ─── */
    function startLockoutCountdown(seconds, btn, inp) {
        var remaining = seconds;
        inp.disabled  = true;
        btn.disabled  = true;

        function tick() {
            var m = Math.floor(remaining / 60);
            var s = remaining % 60;
            var label = m > 0
                ? m + ' menit' + (s ? ' ' + s + ' dtk' : '')
                : s + ' detik';
            setError('Terlalu banyak percobaan. Tunggu ' + label + ' lagi.');
            btn.textContent = label;

            if (remaining <= 0) {
                inp.disabled    = false;
                btn.disabled    = false;
                btn.textContent = 'VERIFIKASI';
                inp.value       = '';
                setError('');
                inp.focus();
                return;
            }
            remaining--;
            setTimeout(tick, 1000);
        }
        tick();
    }

    /* ─── Public API ─── */
    window.requireOTP = function (phone, onSuccess) {
        checkDevice(phone).then(function (trusted) {
            if (trusted) {
                onSuccess();
            } else {
                openModal(phone, onSuccess);
            }
        });
    };

})();
