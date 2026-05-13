<?php
session_start();
require_once __DIR__ . '/../includes/config.php';

if (!empty($_SESSION['adk_admin'])) {
    header('Location: /adk/admin/dashboard.php');
    exit;
}

$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $email    = trim($_POST['email']    ?? '');
    $password = trim($_POST['password'] ?? '');

    if ($email === ADMIN_EMAIL && password_verify($password, ADMIN_PASSWORD_HASH)) {
        session_regenerate_id(true);
        $_SESSION['adk_admin'] = true;
        header('Location: /adk/admin/dashboard.php');
        exit;
    } else {
        $error = 'Login gagal.';
    }
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Admin — ADK</title>
    <style>
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        body {
            font-family: 'Segoe UI', sans-serif;
            background: linear-gradient(135deg, #0e0b28, #1a1043);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            padding: 20px;
        }

        .login-wrap {
            width: 100%;
            max-width: 400px;
        }

        .login-brand {
            text-align: center;
            margin-bottom: 32px;
        }

        .login-brand-title {
            font-size: 20px;
            font-weight: 800;
            color: white;
            letter-spacing: 0.5px;
        }

        .login-brand-sub {
            font-size: 12px;
            color: #6b7280;
            margin-top: 4px;
        }

        .login-card {
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(124,58,237,0.3);
            border-radius: 20px;
            padding: 36px 32px;
        }

        .login-title {
            font-size: 17px;
            font-weight: 800;
            color: white;
            margin-bottom: 24px;
            text-align: center;
        }

        .form-group { margin-bottom: 18px; }

        .form-label {
            display: block;
            font-size: 11px;
            font-weight: 700;
            color: #9ca3af;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            margin-bottom: 8px;
        }

        .form-input {
            width: 100%;
            padding: 12px 16px;
            background: rgba(255,255,255,0.07);
            border: 1px solid rgba(124,58,237,0.35);
            border-radius: 10px;
            color: white;
            font-size: 15px;
            outline: none;
            transition: border-color 0.2s, background 0.2s;
            font-family: inherit;
        }

        .form-input:focus {
            border-color: #7c3aed;
            background: rgba(124,58,237,0.12);
        }

        .form-input::placeholder { color: rgba(255,255,255,0.25); }

        .error-msg {
            background: rgba(239,68,68,0.1);
            border: 1px solid rgba(239,68,68,0.3);
            border-radius: 10px;
            padding: 10px 14px;
            color: #f87171;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 18px;
            text-align: center;
        }

        .btn-login {
            width: 100%;
            padding: 13px;
            background: #7c3aed;
            color: white;
            border: none;
            border-radius: 12px;
            font-size: 15px;
            font-weight: 700;
            cursor: pointer;
            letter-spacing: 0.5px;
            transition: background 0.2s, box-shadow 0.2s;
            font-family: inherit;
            margin-top: 6px;
        }

        .btn-login:hover {
            background: #a855f7;
            box-shadow: 0 6px 20px rgba(124,58,237,0.45);
        }
    </style>
</head>
<body>
    <div class="login-wrap">
        <div class="login-brand">
            <div class="login-brand-title">ADK PHOTOCOPY</div>
            <div class="login-brand-sub">Panel Admin</div>
        </div>

        <div class="login-card">
            <div class="login-title">Masuk ke Dashboard</div>

            <?php if ($error): ?>
            <div class="error-msg"><?= htmlspecialchars($error) ?></div>
            <?php endif; ?>

            <form method="POST" autocomplete="off">
                <div class="form-group">
                    <label class="form-label" for="email">Email</label>
                    <input
                        type="email"
                        id="email"
                        name="email"
                        class="form-input"
                        placeholder="admin@email.com"
                        required
                        autofocus
                    >
                </div>
                <div class="form-group">
                    <label class="form-label" for="password">Password</label>
                    <input
                        type="password"
                        id="password"
                        name="password"
                        class="form-input"
                        placeholder="••••••••"
                        required
                    >
                </div>
                <button type="submit" class="btn-login">MASUK</button>
            </form>
        </div>
    </div>
</body>
</html>
