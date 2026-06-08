<?php
session_start();
$cfg = require __DIR__ . '/config/clave.php';
$error = '';
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $u = $_POST['usuario'] ?? '';
    $c = $_POST['clave']   ?? '';
    if ($u === $cfg['usuario'] && $c === $cfg['clave']) {
        $_SESSION['logueado'] = true;
        $_SESSION['usuario']  = $u;
        header('Location: INICIO.php'); exit;
    }
    $error = 'Usuario o clave incorrectos.';
}
?>
<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8">
<title>Ingreso - Griff Salud</title>
<style>
  :root{--azul:#1A2D9C;--celeste:#29ABE2;--bg:#eef2f9;--card:#fff;--line:#e2e7f0;--txt:#15246E;--muted:#6b7793;}
  *{margin:0;padding:0;box-sizing:border-box;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;}
  body{background:var(--bg);min-height:100vh;display:flex;align-items:center;justify-content:center;color:var(--txt);}
  .box{background:var(--card);border:1px solid var(--line);border-top:4px solid var(--celeste);
    border-radius:13px;padding:32px 36px;width:340px;box-shadow:0 4px 20px rgba(26,45,156,.08);}
  .box img{height:46px;display:block;margin:0 auto 14px;}
  h1{font-size:17px;font-weight:700;color:var(--azul);text-align:center;margin-bottom:6px;}
  .sub{font-size:12px;color:var(--muted);text-align:center;margin-bottom:22px;}
  label{display:block;font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;font-weight:600;margin-bottom:5px;}
  input{width:100%;border:1px solid var(--line);border-radius:7px;padding:9px 11px;font-size:14px;color:var(--txt);}
  input:focus{outline:none;border-color:var(--celeste);}
  .fld{margin-bottom:14px;}
  button{width:100%;background:var(--azul);color:#fff;border:none;border-radius:7px;padding:11px;
    font-size:14px;font-weight:600;cursor:pointer;margin-top:6px;}
  button:hover{background:#15247a;}
  .err{background:#FFE0E0;color:#b91c1c;font-size:13px;padding:8px 11px;border-radius:7px;
    margin-bottom:14px;text-align:center;}
</style></head><body>
<div class="box">
  <img src="assets/griff.png" alt="Griff Salud">
  <h1>Tablero General</h1>
  <div class="sub">Ingresar para continuar</div>
  <?php if ($error): ?><div class="err"><?= htmlspecialchars($error) ?></div><?php endif; ?>
  <form method="post">
    <div class="fld"><label>Usuario</label><input type="text" name="usuario" required autofocus></div>
    <div class="fld"><label>Clave</label><input type="password" name="clave" required></div>
    <button type="submit">Ingresar</button>
  </form>
</div>
</body></html>
