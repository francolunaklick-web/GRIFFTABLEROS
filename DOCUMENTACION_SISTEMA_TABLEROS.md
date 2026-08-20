# Sistema de Tableros Web — Documentación Técnica
**Para replicar en nuevos clientes · Griff Salud como referencia**

---

## 1. QUÉ ES EL SISTEMA Y CÓMO FUNCIONA

Es una plataforma web privada (con login y contraseña) que muestra dashboards interactivos de gestión. Los datos viven en archivos Excel en la computadora. Un script Python lee esos Excel, genera una página web (PHP) con los datos embebidos, y la sube a internet via GitHub + Hostinger.

**Flujo resumido:**

```
Excel (datos) → Python (actualizar.py) → PHP/HTML (dashboard) → GitHub (push) → Hostinger (web pública)
```

El usuario nunca accede directamente a los Excel. Ve la web, con gráficos, tablas y filtros, protegida por contraseña.

---

## 2. STACK TECNOLÓGICO

| Componente | Qué hace |
|---|---|
| **Excel (.xlsx)** | Fuente de datos. El equipo los actualiza manualmente. |
| **Python + openpyxl** | Lee los Excel y genera los archivos PHP con los datos |
| **PHP** | Lenguaje del servidor web. Hostinger lo ejecuta gratis con cualquier plan. |
| **HTML + CSS + JavaScript** | La interfaz visual del dashboard (charts, filtros, KPIs) |
| **Git / GitHub** | Control de versiones y canal de publicación |
| **Hostinger** | Hosting web. Conectado a GitHub: cada `git push` actualiza la web automáticamente |
| **Visual Studio Code** | Editor donde se corren los scripts Python (con terminal integrada) |

---

## 3. ESTRUCTURA DE CARPETAS (referencia: Griff Salud)

```
TABLERO GENERAL - GRIFF WEB/          ← Carpeta raíz = repositorio Git
│
├── .git/                              ← Git (no tocar)
├── .gitignore                         ← Archivos que Git ignora (backups, locks, etc.)
├── .htaccess                          ← Configuración Apache (Hostinger)
│
├── ACTUALIZAR_TODO.bat                ← Corre todos los scripts Python (Windows)
├── ACTUALIZAR_Y_PUBLICAR.sh           ← Corre scripts + hace git push (Linux/Mac)
│
├── 1_DATOS/                           ← FUENTE DE DATOS (Excel e inputs)
│   ├── comercial/
│   │   ├── BBDD_COMERCIAL.xlsx        ← Base de datos de ventas + padrón
│   │   ├── MARKETING.xlsx             ← Gastos Meta Ads por mes
│   │   └── cobros/                    ← CSVs de cobranzas exportados del sistema
│   ├── estado_resultado/              ← Excel mensuales de Estado de Resultado
│   ├── costos_fijos/                  ← Excel de costos fijos por mes
│   ├── prestadores/                   ← Planillas de prestadores por mes
│   ├── cheques/                       ← CSVs de cheques emitidos
│   ├── facturacion/                   ← Excel de facturación e impuestos
│   └── flujo_bancario/                ← CSVs de saldos y movimientos
│
├── 2_TABLEROS/                        ← SCRIPTS Y PROCESAMIENTO
│   ├── TABLERO COMERCIAL/
│   │   ├── 2_scripts/actualizar.py    ← Script principal (lee BBDD → genera dashboard)
│   │   └── 3_outputs/                 ← Snapshots JSON y logs (no se publican directamente)
│   ├── TABLERO ESTADO DE RESULTADO/
│   │   ├── 2_scripts/actualizar.py
│   │   ├── 2_scripts/lib/             ← Módulos Python auxiliares
│   │   └── 3_outputs/dashboard.php    ← Output generado
│   ├── TABLERO CUENTAS A PAGAR CLOUD/
│   ├── TABLERO FLUJO BANCARIO/
│   ├── TABLERO FACTURACION/
│   └── TABLERO PROYECCION DE PAGOS/
│
├── 3_RESULTADOS/                      ← (Carpeta auxiliar, no crítica)
│
├── assets/
│   └── griff.png                      ← Logo
│
├── config/
│   ├── clave.php                      ← Contraseña hasheada (no se sube a GitHub)
│   └── .htaccess                      ← Bloquea acceso web a esta carpeta
│
│   ── CARPETAS PÚBLICAS (las que Hostinger sirve) ──
│
├── comercial/
│   └── dashboard.php                  ← Tablero Comercial (generado por actualizar.py)
├── estado_resultado/
│   └── dashboard.php
├── cuentas_a_pagar/
│   └── tablero.php
├── flujo_bancario/
│   └── tablero.php
├── facturacion/
│   └── dashboard.php
├── proyeccion/
│   └── proyeccion.php
│
├── INICIO.php                         ← Menú principal con links a todos los tableros
├── index.php                          ← Redirige a INICIO.php (o al login si no está logueado)
├── login.php                          ← Página de login con contraseña
├── logout.php                         ← Cierra sesión
└── auth_check.php                     ← Se incluye en cada tablero para verificar el login
```

---

## 4. CÓMO FUNCIONA CADA SCRIPT (actualizar.py)

Cada tablero tiene su propio `actualizar.py`. El patrón es siempre el mismo:

```
1. Lee los archivos Excel de 1_DATOS/ usando openpyxl
2. Procesa los datos (suma, filtra, calcula KPIs)
3. Genera un string PHP/HTML con los datos embebidos como JSON
4. Escribe ese string en el archivo de salida (dashboard.php, tablero.php)
```

Ejemplo simplificado:

```python
import openpyxl, json

# 1. Leer Excel
wb = openpyxl.load_workbook("datos.xlsx", data_only=True)
ws = wb.active
datos = [{"nombre": r[0], "valor": r[1]} for r in ws.iter_rows(min_row=2, values_only=True)]

# 2. Calcular
total = sum(d["valor"] for d in datos)

# 3. Generar PHP
php = f"""<?php require_once '../auth_check.php'; ?>
<!DOCTYPE html>
<html>
<body>
<script>
const DATOS = {json.dumps(datos)};
const TOTAL = {total};
</script>
<!-- HTML del dashboard aquí -->
</body>
</html>"""

# 4. Guardar
with open("../../mi_tablero/dashboard.php", "w", encoding="utf-8") as f:
    f.write(php)
```

Los datos quedan "congelados" en el PHP. No hay base de datos, no hay consultas en tiempo real. Cada vez que el usuario corre el script, el PHP se regenera con los datos actuales.

---

## 5. SISTEMA DE LOGIN Y AUTENTICACIÓN

Es una autenticación PHP simple de contraseña única (una clave para todos).

**Archivos clave:**

`config/clave.php` — solo contiene el hash de la contraseña:
```php
<?php
$CLAVE_HASH = '$2y$10$xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
```

Para generar el hash de una contraseña nueva, ejecutar en PHP:
```php
echo password_hash("mi_contraseña_nueva", PASSWORD_DEFAULT);
```

`login.php` — formulario que verifica con `password_verify()` y guarda `$_SESSION['logueado'] = true`.

`auth_check.php` — se incluye al inicio de cada tablero. Si no hay sesión activa, redirige al login.

`config/.htaccess` — bloquea el acceso web directo a la carpeta config (la contraseña no es visible desde el navegador).

---

## 6. GITHUB — CONFIGURACIÓN

**Repositorio actual de Griff:** `https://github.com/francolunaklick-web/GRIFFTABLEROS`

### Para un cliente nuevo:

1. Crear cuenta en GitHub (o usar la existente)
2. Crear repositorio nuevo (privado recomendado)
3. En la carpeta del proyecto, inicializar git:
   ```bash
   git init
   git remote add origin https://github.com/tu-usuario/nombre-repo.git
   git add -A
   git commit -m "Primer commit"
   git push -u origin main
   ```

### El .gitignore importante:
```
config/clave.php        # La contraseña NUNCA sube a GitHub
1_DATOS/**/*.xlsx       # Los Excel tampoco (datos privados del cliente)
**/__pycache__/
**/*.pyc
**/.~lock.*             # Archivos de bloqueo de LibreOffice
```

---

## 7. HOSTINGER — CONFIGURACIÓN

Hostinger sirve los archivos PHP como web pública. Se conecta a GitHub para despliegue automático.

### Pasos para un cliente nuevo:

**1. Crear hosting en Hostinger**
- Plan Business (permite PHP, Git auto-deploy)
- Asignar dominio (el del cliente o un subdominio como `tableros.cliente.com`)

**2. Conectar con GitHub (Auto Deploy)**
- En el Panel de Hostinger → Git → conectar repositorio
- Branch: `main`
- Directorio de deploy: `public_html/` (o la carpeta raíz del dominio)
- Cada vez que hagas `git push`, Hostinger actualiza la web automáticamente

**3. Habilitar PHP y configurar .htaccess**
El `.htaccess` en la raíz maneja las redirecciones. El que está en `config/` bloquea el acceso externo:
```apache
# config/.htaccess
Deny from all
```

**4. Generar la clave hasheada en Hostinger**
En el File Manager de Hostinger o via SSH, correr:
```php
php -r "echo password_hash('clave_del_cliente', PASSWORD_DEFAULT);"
```
Y pegar el resultado en `config/clave.php`.

---

## 8. FLUJO COMPLETO DE ACTUALIZACIÓN

Cuando hay datos nuevos, el proceso es:

```
1. El equipo actualiza los Excel en 1_DATOS/
2. Franco abre Visual Studio Code
3. Corre el script: python 2_scripts/actualizar.py  (desde la terminal de VS Code)
   - O para todos juntos: corre ACTUALIZAR_TODO.bat
4. El script genera/actualiza los archivos .php en las carpetas públicas
5. Franco hace git push (manual o via ACTUALIZAR_Y_PUBLICAR.sh)
6. Hostinger detecta el push y actualiza la web en segundos
7. El cliente entra a la web y ve los datos actualizados
```

**IMPORTANTE:** Solo los archivos PHP y auxiliares (CSV, JSON) se suben a GitHub y llegan a la web. Los Excel con datos del cliente NUNCA se publican.

---

## 9. TABLEROS EXISTENTES EN GRIFF

| Tablero | Script | Output público | Datos fuente |
|---|---|---|---|
| Comercial | `TABLERO COMERCIAL/2_scripts/actualizar.py` | `comercial/dashboard.php` | `BBDD_COMERCIAL.xlsx`, `MARKETING.xlsx`, cobros CSV |
| Estado de Resultado | `TABLERO ESTADO DE RESULTADO/2_scripts/actualizar.py` | `estado_resultado/dashboard.php` | Excel mensuales de ER |
| Cuentas a Pagar | `TABLERO CUENTAS A PAGAR CLOUD/2_scripts/actualizar.py` | `cuentas_a_pagar/tablero.php` | Planillas prestadores, OPs, cheques |
| Flujo Bancario | `TABLERO FLUJO BANCARIO/2_scripts/actualizar.py` | `flujo_bancario/tablero.php` | Saldos CSV, cheques CSV, movimientos |
| Facturación | `TABLERO FACTURACION/actualizar.py` | `facturacion/dashboard.php` | Excel facturación |
| Proyección de Pagos | `TABLERO PROYECCION DE PAGOS/2_scripts/actualizar.py` | `proyeccion/proyeccion.php` | Excel proyección |

---

## 10. CÓMO REPLICAR PARA UN CLIENTE NUEVO

### Opción A — Una web por cliente (recomendado para empezar)

Cada cliente tiene su propio hosting, dominio y repositorio. Más simple, más aislado.

**Pasos:**
1. Copiar toda la carpeta `TABLERO GENERAL - GRIFF WEB/` y renombrarla para el cliente
2. Reemplazar logo (`assets/`) y paleta de colores en los PHP/CSS
3. Vaciar los Excel de datos de Griff, armar los del cliente con la misma estructura de columnas
4. Ajustar los scripts Python si el cliente tiene datos distintos (columnas, categorías, etc.)
5. Crear repo GitHub nuevo → conectar a Hostinger nuevo → primer push
6. Configurar la contraseña del cliente en `config/clave.php`

**Costo estimado por cliente:** ~$5-10 USD/mes hosting Hostinger + dominio (~$10 USD/año)

### Opción B — Una sola web, múltiples clientes (más complejo, más escalable)

Una web con login por usuario. Cada usuario solo ve sus tableros.

**Cambios necesarios respecto al sistema actual:**

1. Reemplazar la clave única por usuarios con base de datos:
   - Agregar MySQL en Hostinger (viene incluido en planes Business)
   - Tabla `usuarios`: `id, nombre, empresa, password_hash, tableros_acceso`
   - `login.php` verifica usuario+contraseña y guarda `$_SESSION['empresa']`

2. Los tableros se generan con nombres únicos por cliente:
   - `comercial_acmecorp/dashboard.php`
   - `comercial_otroempresa/dashboard.php`

3. `auth_check.php` verifica que el usuario tenga acceso al tablero que intenta ver

4. El `INICIO.php` muestra solo los tableros habilitados para el usuario logueado

**Ventaja:** Un solo hosting sirve a todos. **Desventaja:** Más código que mantener, riesgo de que un cliente vea datos del otro si hay un bug.

---

## 11. DEPENDENCIAS PYTHON

Cada tablero usa openpyxl como mínimo. Instalar con:

```bash
pip install openpyxl
```

Algunos tableros más complejos tienen `requirements.txt`. Instalar con:

```bash
pip install -r requirements.txt
```

En Windows, VS Code debe tener Python instalado y en el PATH. Verificar con:
```bash
python --version
```

---

## 12. CHECKLIST PARA NUEVA INSTANCIA

- [ ] Copiar estructura de carpetas del proyecto Griff
- [ ] Crear repositorio GitHub (privado)
- [ ] Agregar `.gitignore` con `config/clave.php` y los Excel
- [ ] Crear hosting Hostinger y conectar GitHub (Auto Deploy)
- [ ] Configurar dominio o subdominio del cliente
- [ ] Generar hash de contraseña y guardar en `config/clave.php`
- [ ] Armar los Excel de datos con la estructura correcta para cada tablero
- [ ] Ajustar los scripts Python según los datos del cliente
- [ ] Primer `git push` → verificar que Hostinger deploye bien
- [ ] Verificar login en la web del cliente
- [ ] Verificar que cada tablero cargue correctamente
- [ ] Entregar la URL y la contraseña al cliente

---

## 13. DATOS QUE CLAUDE NECESITA AL INICIO DE CADA SESIÓN

Para que una nueva sesión de Claude entienda el contexto sin explicar todo de cero, proporcionar:

1. **La carpeta del proyecto** (seleccionarla en Cowork)
2. **Qué tablero se va a trabajar** (ej: "tablero comercial")
3. **Qué datos maneja** (ej: ventas por asesora, cápitas, cobros)
4. **La regla crítica:** *"Solo modificar datos y scripts. El usuario corre los scripts desde Visual Studio Code. No ejecutar scripts."*
5. Si hay Excel para revisar, mencionarlo explícitamente (ej: "la BBDD tiene columnas X, Y, Z")

---

*Documento generado en base al sistema Griff Salud · Agosto 2026*

---

# APÉNDICE — CÓDIGO FUENTE COMPLETO

> Estos archivos son suficientes para replicar el sistema desde cero en un cliente nuevo.
> Reemplazar las referencias a "Griff Salud", "griff.png", colores y nombres de asesoras según el cliente.

---

## A1. `auth_check.php`
Incluir al inicio de cada tablero para protegerlo con login.

```php
<?php
// Verificacion de sesion para todas las paginas protegidas.
// Si no esta logueado, redirige a login.php (calcula la URL relativa al root).
if (session_status() === PHP_SESSION_NONE) {
    session_start();
}
if (empty($_SESSION['logueado'])) {
    $rootDir = realpath(__DIR__);
    $self    = realpath($_SERVER['SCRIPT_FILENAME'] ?? __FILE__);
    $depth   = 0;
    if ($self && $rootDir) {
        $rel = str_replace($rootDir, '', $self);
        $depth = substr_count(trim($rel, '/\\'), DIRECTORY_SEPARATOR);
        if ($depth === 0) {
            $depth = substr_count(trim($rel, '/\\'), '/');
        }
    }
    $up = str_repeat('../', max(0, $depth));
    header('Location: ' . $up . 'login.php');
    exit;
}
```

---

## A2. `login.php`
Página de acceso con contraseña única.

```php
<?php
ob_start();
session_start();
require_once __DIR__ . '/config/clave.php';
$error = '';
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $ingresada = $_POST['clave'] ?? '';
    if (password_verify($ingresada, $CLAVE_HASH)) {
        $_SESSION['logueado'] = true;
        header('Location: INICIO.php');
        exit;
    } else {
        $error = 'Contraseña incorrecta.';
    }
}
?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Acceso — Tableros</title>
    <style>
        :root { --azul:#1A2D9C; --celeste:#29ABE2; --bg:#eef2f9; --card:#fff; --line:#e2e7f0; --muted:#6b7793; --azul-txt:#15246E; }
        * { margin:0; padding:0; box-sizing:border-box; }
        body { font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif; background:var(--bg); min-height:100vh; display:flex; align-items:center; justify-content:center; padding:24px; }
        .wrap { width:100%; max-width:420px; }
        .card { background:var(--card); border:1px solid var(--line); border-radius:14px; border-top:4px solid var(--celeste); padding:36px 32px; }
        .head { display:flex; align-items:center; gap:16px; margin-bottom:28px; padding-bottom:24px; border-bottom:1px solid var(--line); }
        .head img { height:52px; width:auto; }
        .head h1 { font-size:17px; font-weight:700; color:var(--azul); }
        .head p { font-size:12px; color:var(--muted); margin-top:3px; }
        label { display:block; font-size:13px; font-weight:600; color:var(--azul-txt); margin-bottom:7px; }
        input[type=password] { width:100%; padding:10px 14px; border:1px solid var(--line); border-radius:8px; font-size:14px; color:var(--azul-txt); background:#f8fafd; outline:none; }
        input[type=password]:focus { border-color:var(--celeste); background:#fff; }
        button { width:100%; margin-top:18px; padding:11px; background:var(--azul); color:white; border:none; border-radius:8px; font-size:14px; font-weight:600; cursor:pointer; }
        button:hover { background:#152380; }
        .error { background:#fff0f0; border:1px solid #fcc; border-radius:8px; padding:10px 14px; font-size:13px; color:#c0392b; margin-bottom:16px; }
        .foot { text-align:center; font-size:11px; color:var(--muted); margin-top:20px; }
    </style>
</head>
<body>
<div class="wrap">
    <div class="card">
        <div class="head">
            <img src="/assets/logo.png" alt="Logo">
            <div>
                <h1>Tableros de Informes</h1>
                <p>Ingresá tu contraseña para acceder</p>
            </div>
        </div>
        <?php if ($error): ?>
            <div class="error"><?= htmlspecialchars($error) ?></div>
        <?php endif; ?>
        <form method="POST">
            <label>Contraseña de acceso</label>
            <input type="password" name="clave" autofocus placeholder="••••••••">
            <button type="submit">Ingresar →</button>
        </form>
    </div>
    <div class="foot">Acceso restringido</div>
</div>
</body>
</html>
```

---

## A3. `logout.php`

```php
<?php
session_start();
session_unset();
session_destroy();
header('Location: login.php');
exit;
```

---

## A4. `index.php`

```php
<?php
session_start();
if (!empty($_SESSION['logueado'])) {
    header('Location: INICIO.php');
} else {
    header('Location: login.php');
}
exit;
```

---

## A5. `config/clave.php`
**NUNCA subir a GitHub** (está en .gitignore).
Para generar un hash nuevo: ir a https://bcrypt.online, poner la clave, Rounds=12.

```php
<?php
// Contraseña actual: [la que definas para el cliente]
$CLAVE_HASH = '$2y$10$REEMPLAZAR_CON_HASH_GENERADO_EN_BCRYPT_ONLINE';
?>
```

---

## A6. `config/.htaccess`
Bloquea el acceso web a la carpeta config.

```apache
Deny from all
```

---

## A7. `.htaccess` (raíz)
El index del sitio arranca en login.php.

```apache
DirectoryIndex login.php

<FilesMatch "\.(csv|xlsx|xls|pdf)$">
    Require all denied
</FilesMatch>
```

---

## A8. `.gitignore`
Qué NO sube a GitHub.

```
1_DATOS/
2_TABLEROS/
3_RESULTADOS/
9_archivo/
ACTUALIZAR_TODO.bat
0_LEEME.txt
config/clave.php
**/__pycache__/
**/*.pyc
**/.~lock.*
```

> **Importante:** Agregar `config/clave.php` al .gitignore del proyecto Griff si aún no está.

---

## A9. `ACTUALIZAR_TODO.bat`
Corre todos los scripts Python desde Windows (doble clic).

```bat
@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo    ACTUALIZAR TODOS LOS TABLEROS
echo ============================================================

echo [1] Cuentas a Pagar...
cd /d "%~dp0\2_TABLEROS\TABLERO CUENTAS A PAGAR CLOUD"
python "2_scripts\actualizar.py"

echo [2] Estado de Resultado...
cd /d "%~dp0\2_TABLEROS\TABLERO ESTADO DE RESULTADO"
python "2_scripts\actualizar.py"

echo [3] Facturacion...
cd /d "%~dp0\2_TABLEROS\TABLERO FACTURACION"
python "actualizar.py"

echo [4] Flujo Bancario...
cd /d "%~dp0\2_TABLEROS\TABLERO FLUJO BANCARIO"
python "2_scripts\actualizar.py"

echo [5] Proyeccion de Pagos...
cd /d "%~dp0\2_TABLEROS\TABLERO PROYECCION DE PAGOS"
python "2_scripts\actualizar.py"

echo [6] Comercial...
cd /d "%~dp0\2_TABLEROS\TABLERO COMERCIAL"
python "2_scripts\actualizar.py"

echo ============================================================
echo    LISTO
echo ============================================================
pause
```

---

## A10. `ACTUALIZAR_Y_PUBLICAR.sh`
Corre scripts + copia outputs a carpetas públicas + hace git push.
Usar en Mac/Linux o WSL en Windows.

```bash
#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

echo "=== 1) ACTUALIZAR TABLEROS ==="
( cd "2_TABLEROS/TABLERO CUENTAS A PAGAR CLOUD" && python 2_scripts/actualizar.py )
( cd "2_TABLEROS/TABLERO ESTADO DE RESULTADO"   && python 2_scripts/actualizar.py )
( cd "2_TABLEROS/TABLERO FACTURACION"           && python actualizar.py )
( cd "2_TABLEROS/TABLERO FLUJO BANCARIO"        && python 2_scripts/actualizar.py )
( cd "2_TABLEROS/TABLERO PROYECCION DE PAGOS"   && python 2_scripts/actualizar.py )

echo "=== 2) PUBLICAR ==="

publish() {
  local SRC_FILE="$1" DEST_FILE="$2"
  if [ -f "$SRC_FILE" ]; then
    { echo '<?php require_once __DIR__ . "/../auth_check.php"; ?>'
      sed -e '1{/<?php.*auth_check.*?>/d;}' "$SRC_FILE"
    } > "$DEST_FILE"
  fi
}

mkdir -p cuentas_a_pagar estado_resultado facturacion flujo_bancario proyeccion

cp -f "2_TABLEROS/TABLERO CUENTAS A PAGAR CLOUD/3_outputs"/*.csv cuentas_a_pagar/ 2>/dev/null || true
publish "2_TABLEROS/TABLERO CUENTAS A PAGAR CLOUD/3_outputs/tablero.php" cuentas_a_pagar/tablero.php

cp -f "2_TABLEROS/TABLERO ESTADO DE RESULTADO/3_outputs"/*.xlsx estado_resultado/ 2>/dev/null || true
publish "2_TABLEROS/TABLERO ESTADO DE RESULTADO/3_outputs/dashboard.php" estado_resultado/dashboard.php

publish "2_TABLEROS/TABLERO FACTURACION/dashboard.php" facturacion/dashboard.php

cp -f "2_TABLEROS/TABLERO FLUJO BANCARIO/3_outputs"/*.csv flujo_bancario/ 2>/dev/null || true
publish "2_TABLEROS/TABLERO FLUJO BANCARIO/3_outputs/tablero.php" flujo_bancario/tablero.php

publish "2_TABLEROS/TABLERO PROYECCION DE PAGOS/2_scripts/plantilla.php" proyeccion/proyeccion.php

echo "=== 3) GIT PUSH ==="
git add -A
git -c user.email="tablero@cliente.com" -c user.name="Tablero Auto" \
    commit -m "Actualizacion automatica $(date +'%Y-%m-%d %H:%M')" || echo "(sin cambios)"
git push origin main || echo "ATENCION: revisar remoto"

echo "=== LISTO ==="
```

---

## A11. `INICIO.php`
Menú principal con links a todos los tableros.

```php
<?php require_once __DIR__ . '/auth_check.php'; ?>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Tableros de Informes</title>
<style>
  :root { --azul:#1A2D9C; --azul-txt:#15246E; --celeste:#29ABE2; --bg:#eef2f9; --card:#fff; --line:#e2e7f0; --muted:#6b7793; }
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif; background:var(--bg); color:var(--azul-txt); padding:36px 24px; min-height:100vh; }
  .wrap { max-width:980px; margin:0 auto; }
  .head { background:var(--card); border-radius:14px; border:1px solid var(--line); padding:22px 26px; display:flex; align-items:center; gap:20px; border-top:4px solid var(--celeste); }
  .head img { height:58px; width:auto; }
  .head h1 { font-size:20px; font-weight:700; color:var(--azul); }
  .head .sub { font-size:13px; color:var(--muted); margin-top:3px; }
  .grid { display:grid; grid-template-columns:repeat(2,1fr); gap:16px; margin-top:20px; }
  .card { display:block; text-decoration:none; color:inherit; background:var(--card); border:1px solid var(--line); border-radius:14px; padding:22px 24px; transition:transform .14s,box-shadow .14s,border-color .14s; }
  .card:hover { transform:translateY(-3px); box-shadow:0 8px 22px rgba(26,45,156,.13); border-color:var(--celeste); }
  .cname { font-size:17px; font-weight:700; color:var(--azul); }
  .cdesc { font-size:13px; color:#56607a; line-height:1.6; margin-top:8px; }
  .clink { font-size:12.5px; font-weight:700; margin-top:14px; }
  .foot { text-align:center; font-size:12px; color:var(--muted); margin-top:24px; padding-top:16px; border-top:1px solid var(--line); }
  @media(max-width:640px) { .grid { grid-template-columns:1fr; } }
</style>
</head>
<body>
<div class="wrap">
  <div class="head">
    <img src="assets/logo.png" alt="Logo">
    <div>
      <h1>Tableros de Informes</h1>
      <div class="sub">Seleccioná un tablero para abrirlo</div>
    </div>
  </div>
  <div class="grid">
    <!-- Agregar/quitar cards según los tableros del cliente -->
    <a class="card" href="comercial/dashboard.php">
      <div class="cname">Comercial</div>
      <div class="cdesc">Embudo de ventas, KPIs por asesora, conversión y seguimiento.</div>
      <div class="clink" style="color:#1E8449;">Abrir tablero →</div>
    </a>
    <a class="card" href="estado_resultado/dashboard.php">
      <div class="cname">Estado de Resultado</div>
      <div class="cdesc">Ingresos vs costos, resultado del mes.</div>
      <div class="clink" style="color:#29ABE2;">Abrir tablero →</div>
    </a>
  </div>
  <div class="foot">Para refrescar los datos: corré <b>ACTUALIZAR_TODO.bat</b></div>
</div>
</body>
</html>
```

---

## A12. `2_TABLEROS/TABLERO COMERCIAL/2_scripts/actualizar.py`
Script completo del Tablero Comercial. Es el más complejo y sirve de referencia para construir otros.

```python
# -*- coding: utf-8 -*-
"""actualizar.py — Tablero Comercial"""

import os, sys, json, shutil, re
from datetime import datetime, date

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
ROOT         = os.path.dirname(SCRIPT_DIR)
HUB          = os.path.dirname(os.path.dirname(ROOT))

BBDD_PATH    = os.path.join(HUB, "1_DATOS", "comercial", "BBDD_COMERCIAL.xlsx")
OUT_PHP      = os.path.join(HUB, "comercial", "dashboard.php")
SNAPSHOT_OLD  = os.path.join(ROOT, "3_outputs", "snapshot_anterior.json")
SNAPSHOT_NEW  = os.path.join(ROOT, "3_outputs", "snapshot_actual.json")
KPI_HISTORICO = os.path.join(ROOT, "3_outputs", "kpi_historico.json")

os.makedirs(os.path.join(HUB, "comercial"), exist_ok=True)
os.makedirs(os.path.join(ROOT, "3_outputs"), exist_ok=True)

# ── escalas comision ─────────────────────────────────────────────────────────
ESCALA_FUERA  = [(40,400000,5714.29),(52,520000,5714.29),(60,780000,7428.57),(68,1100000,9243.70)]
ESCALA_DENTRO = [(40,340000,4857.14),(52,442000,4857.14),(60,663000,6314.29),(68, 935000,7857.14)]

def get_tramo(capitas, escala):
    t = None
    for mn, com, vpc in escala:
        if capitas >= mn: t = {"min":mn,"comision":com,"vpc":vpc}
    return t

def grupo_a_capitas(grupo):
    if not grupo: return 0
    g = str(grupo).strip().lower().replace(' ','')
    if g=='individual': return 1
    if g=='individual+hijo': return 2
    if g=='+pareja': return 2
    if '+pareja+1hijo' in g: return 3
    if '+pareja+2hijos' in g: return 4
    if '+pareja+3hijos' in g: return 5
    if 'individual+1adherente' in g: return 2
    if 'individual+2adherentes' in g: return 3
    m = re.search(r'\((\d+)\)', g)
    if m: return int(m.group(1))
    return 1

def ars(n):
    if not n: return "$ 0"
    return "$ {:,.0f}".format(float(n)).replace(",","X").replace(".",",").replace("X",".")

# ── tabla de precios ──────────────────────────────────────────────────────────
PRECIO_BASE = {"GS100": 135000, "GS200": 220000}
PRECIOS = {
    ("GS100","Prepago","Individual"):           135000,
    ("GS100","Prepago","Individual + hijo"):    202500,
    ("GS100","Prepago","+ pareja"):             256500,
    ("GS100","Prepago","+ pareja + 1 hijo"):   283500,
    ("GS100","Prepago","+ pareja + 2 hijos"):  310500,
    ("GS100","Prepago","+ pareja + 3 hijos"):  337500,
    ("GS100","Prepago","+ pareja + 4 hijos"):  364500,
    ("GS100","Desregulado","Individual"):        65000,
    ("GS100","Desregulado","Individual + hijo"):132500,
    ("GS100","Desregulado","+ pareja"):         116500,
    ("GS100","Desregulado","+ pareja + 1 hijo"):143500,
    ("GS100","Desregulado","+ pareja + 2 hijos"):170500,
    ("GS100","Desregulado","+ pareja + 3 hijos"):197500,
    ("GS100","Desregulado","+ pareja + 4 hijos"):224500,
    ("GS200","Prepago","Individual"):           220000,
    ("GS200","Prepago","Individual + hijo"):    330000,
    ("GS200","Prepago","+ pareja"):             418000,
    ("GS200","Prepago","+ pareja + 1 hijo"):   462000,
    ("GS200","Prepago","+ pareja + 2 hijos"):  506000,
    ("GS200","Prepago","+ pareja + 3 hijos"):  550000,
    ("GS200","Prepago","+ pareja + 4 hijos"):  594000,
    ("GS200","Desregulado","Individual"):       150000,
    ("GS200","Desregulado","Individual + hijo"):260000,
    ("GS200","Desregulado","+ pareja"):         278000,
    ("GS200","Desregulado","+ pareja + 1 hijo"):322000,
    ("GS200","Desregulado","+ pareja + 2 hijos"):366000,
    ("GS200","Desregulado","+ pareja + 3 hijos"):410000,
    ("GS200","Desregulado","+ pareja + 4 hijos"):454000,
}

def calcular_cuota(plan, tipo, grupo):
    p = str(plan  or "").strip()
    t = str(tipo  or "Prepago").strip()
    g = str(grupo or "").strip()
    if not p or not g: return 0.0
    return float(PRECIOS.get((p, t, g), PRECIO_BASE.get(p, 0)))

# ── leer Excel ───────────────────────────────────────────────────────────────
try: import openpyxl
except ImportError: print("FALTA openpyxl — correr: pip install openpyxl"); sys.exit(1)

print(f"Leyendo {BBDD_PATH}...")
if not os.path.exists(BBDD_PATH): print("ERROR: no existe"); sys.exit(1)

wb = openpyxl.load_workbook(BBDD_PATH, data_only=True)
ws = wb["BBDD"]

col_norm = {}
for c in range(1, ws.max_column+1):
    h = ws.cell(2,c).value
    if h: col_norm[h.replace("\n","").strip().upper()] = c-1

def gc(row, name):
    idx = col_norm.get(name.upper())
    return row[idx] if idx is not None and idx < len(row) else None

contactos = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not any(v for v in r): continue
    fecha_c = gc(r,"FECHACONTACTO")
    if isinstance(fecha_c,(datetime,date)):
        mes = fecha_c.strftime("%m-%Y"); fecha_str = fecha_c.strftime("%d/%m/%Y")
    else:
        mes = ""; fecha_str = ""
    fecha_a = gc(r,"FECHAALTA")
    fecha_alta_str = fecha_a.strftime("%d/%m/%Y") if isinstance(fecha_a,(datetime,date)) else ""
    cuota_raw = gc(r,"CUOTAMENSUAL")
    try: cuota = float(cuota_raw) if cuota_raw else 0.0
    except: cuota = 0.0
    if cuota == 0.0:
        cuota = calcular_cuota(gc(r,"PLANVENDIDO"), gc(r,"TIPO"), gc(r,"GRUPOFAMILIAR"))
    cap_xlsx = gc(r, "CÁPITAS VENTA")
    try: cap = int(float(cap_xlsx)) if cap_xlsx else 0
    except: cap = 0
    if cap == 0:
        cap = grupo_a_capitas(gc(r,"GRUPOFAMILIAR"))
    canal = gc(r,"CANAL") or ""
    contactos.append({
        "nombre":    gc(r,"NOMBRE") or "",
        "telefono":  gc(r,"TELEFONO") or "",
        "canal":     canal,
        "comercial": gc(r,"COMERCIAL") or "",
        "lugar":     gc(r,"LUGAR") or "",
        "estado":    gc(r,"ESTADO") or "",
        "plan":      gc(r,"PLANVENDIDO") or "",
        "tipo":      gc(r,"TIPO") or "",
        "grupo":     gc(r,"GRUPOFAMILIAR") or "",
        "cuota":     cuota,
        "capitas":   cap,
        "mes":       mes,
        "fecha_contacto": fecha_str,
        "fecha_alta": fecha_alta_str,
        "cuil":      gc(r,"CUIL") or "",
        "obs":       gc(r,"OBSERVACIONES") or "",
        "motivo":    gc(r,"MOTIVO_DESCARTE") or "",
        "padron":    "Fuera" if canal=="TERRENO" else "Dentro",
    })

print(f"  {len(contactos)} contactos leidos")

# ── leer hoja MIEMBROS ────────────────────────────────────────────────────────
miembros = []
if "MIEMBROS" in wb.sheetnames:
    ws_miem = wb["MIEMBROS"]
    hs_m = {}
    for c in range(1, ws_miem.max_column+1):
        v = ws_miem.cell(2,c).value
        if v: hs_m[str(v).strip().upper()] = c
    for r in ws_miem.iter_rows(min_row=3, values_only=True):
        if not any(v for v in r): continue
        def gm(idx): return str(r[idx-1] or '').strip() if idx else ''
        miembros.append({
            "cuil_titular": gm(hs_m.get('CUIL_TITULAR',1)),
            "cuil_propio":  gm(hs_m.get('CUIL_PROPIO',2)),
            "nombre":       gm(hs_m.get('NOMBRE',3)),
            "rol":          gm(hs_m.get('ROL',4)),
            "plan":         gm(hs_m.get('PLAN',6)),
        })
    print(f"  {len(miembros)} miembros en hoja MIEMBROS")

ESTADOS     = ["Lead","En Contacto","Cotizado","Cerrado","Descartado"]
COMERCIALES = ["Sol","Vane","Juli"]   # <-- CAMBIAR POR LOS NOMBRES DEL CLIENTE

# ── leer MARKETING.xlsx ───────────────────────────────────────────────────────
import csv, glob
MARKETING_PATH = os.path.join(HUB, "1_DATOS", "comercial", "MARKETING.xlsx")
marketing_data = []
if os.path.exists(MARKETING_PATH):
    wb_mkt = openpyxl.load_workbook(MARKETING_PATH, data_only=True)
    for r in wb_mkt.active.iter_rows(min_row=5, values_only=True):
        if not r[0]: continue
        try:
            marketing_data.append({
                "mes": str(r[0]).strip(), "gasto": float(r[1] or 0),
                "alcance": int(r[2] or 0), "impresiones": int(r[3] or 0),
                "frecuencia": float(r[4] or 0),
                "leads": int(r[5] or 0) if r[5] else 0,
                "notas": str(r[6] or ""),
            })
        except: pass
    print(f"  {len(marketing_data)} meses marketing cargados")

# ── leer Padrón ───────────────────────────────────────────────────────────────
padron = []
for sheet_name in ["PADRON CORDOBA", "PADRON BAHIA"]:
    if sheet_name not in wb.sheetnames: continue
    ws_pad = wb[sheet_name]
    hs_pad = [ws_pad.cell(2,c).value for c in range(1, ws_pad.max_column+1)]
    col_pad = {str(h or '').strip().upper().replace(" ","_"): i for i, h in enumerate(hs_pad) if h}
    for r in ws_pad.iter_rows(min_row=3, values_only=True):
        if not r[0]: continue
        plan = str(r[col_pad.get('PLAN', 3)] or '').strip()
        if plan == 'GSRES': continue
        tipo  = str(r[col_pad.get('TIPO_DE_ALTA', 5)] or 'HISTORICO').strip() or 'HISTORICO'
        rol   = str(r[col_pad.get('ROL', 7)]           or 'Titular').strip()   or 'Titular'
        cuil_p = str(r[col_pad.get('CUIL_PROPIO', 1)]  or '').strip()
        cuil_t = str(r[col_pad.get('CUIL_TITULAR', 6)] or cuil_p).strip()
        dni    = str(r[col_pad.get('DNI', 2)]           or '').strip()
        deleg  = str(r[col_pad.get('DELEGACIÓN', 4)]   or sheet_name).strip()
        padron.append({
            "nombre": str(r[0] or '').strip(), "cuil": cuil_p, "cuil_titular": cuil_t,
            "dni": dni, "plan": plan, "delegacion": deleg, "tipo": tipo, "rol": rol,
            "cuota_est": PRECIO_BASE.get(plan, 0),
        })
print(f"  {len(padron)} afiliados en padrón")

# ── leer Cobros CSV ───────────────────────────────────────────────────────────
COBROS_DIR = os.path.join(HUB, "1_DATOS", "comercial", "cobros")
cobros_por_mes = {}; cobros_resumen = {}; meses_cobros = []
if os.path.exists(COBROS_DIR):
    for csv_path in sorted(glob.glob(os.path.join(COBROS_DIR, "Cobranzas_*.csv"))):
        mes_key = os.path.basename(csv_path).replace("Cobranzas_","").replace(".csv","")
        cobros_mes = {}; total_bruto = 0.0; total_neto = 0.0; count = 0
        try:
            with open(csv_path, encoding='latin-1') as f:
                for row in csv.DictReader(f, delimiter=';'):
                    try:
                        ref  = str(row.get('Referencia','')).strip()
                        imp  = float(str(row.get('Importe_Pagado','0')).replace(',','.') or 0)
                        neto = float(str(row.get('Total_Neto','0')).replace(',','.') or 0)
                        medio= str(row.get('Medio_de_Pago','')).strip()
                        nombre_c = str(row.get('Nombre','')).strip()
                        if ref: cobros_mes[ref] = {"nombre":nombre_c,"importe":imp,"neto":neto,"medio":medio}
                        total_bruto += imp; total_neto += neto; count += 1
                    except: pass
        except Exception as e: print(f"  Error {mes_key}: {e}")
        cobros_por_mes[mes_key] = cobros_mes
        cobros_resumen[mes_key] = {"total_bruto":total_bruto,"total_neto":total_neto,"count":count}
        meses_cobros.append(mes_key)

# ── cruce padrón vs cobros ────────────────────────────────────────────────────
def cruce_padron_cobros(mes_key):
    cobros = cobros_por_mes.get(mes_key, {})
    proy_total = 0; cob_total = 0; resultado = []
    for af in padron:
        if af.get("rol","Titular") == "Adherente": continue
        bonificado = af["tipo"].upper() == "PLAN EMPLEADO"
        dni = af["dni"].lstrip("0"); pago = None
        for ref_key, pag_data in cobros.items():
            if ref_key.lstrip("0") == dni: pago = pag_data; break
        proy = 0 if bonificado else af["cuota_est"]
        cob  = pago["importe"] if pago else 0
        proy_total += proy; cob_total += cob
        resultado.append({**af, "pago":pago, "cobrado":cob,
            "diferencia": 0 if bonificado else (cob - proy),
            "pagado": True if bonificado else pago is not None,
            "bonificado": bonificado})
    return resultado, proy_total, cob_total

cobros_tabla_meses = []
for mk in meses_cobros:
    res, proy, cob = cruce_padron_cobros(mk)
    r_sum = cobros_resumen[mk]
    cobros_tabla_meses.append({"mes":mk,"proyectado":proy,
        "cobrado_bruto":r_sum["total_bruto"],"cobrado_neto":r_sum["total_neto"],
        "count":r_sum["count"],"diferencia":r_sum["total_bruto"]-proy,
        "pct":round(r_sum["total_bruto"]/proy*100,1) if proy else 0})

ultimo_mes_cobros = meses_cobros[-1] if meses_cobros else None
detalle_cobros, proy_ult, cob_ult = cruce_padron_cobros(ultimo_mes_cobros) if ultimo_mes_cobros else ([], 0, 0)
bonif_padron = sum(1 for a in detalle_cobros if a.get("bonificado"))
pagaron      = sum(1 for a in detalle_cobros if a["pagado"] and not a.get("bonificado"))
no_pagaron   = sum(1 for a in detalle_cobros if not a["pagado"] and not a.get("bonificado"))
proy_padron  = sum(a["cuota_est"] for a in padron
               if a["tipo"].upper() != "PLAN EMPLEADO" and a.get("rol","Titular") != "Adherente")

# ── changelog ─────────────────────────────────────────────────────────────────
snap_nuevo = {"fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
    "contactos": [{"nombre":c["nombre"],"telefono":c["telefono"],
                   "estado":c["estado"],"plan":c["plan"],"mes":c["mes"]} for c in contactos]}
cambios = []; fecha_ant = "—"
if os.path.exists(SNAPSHOT_OLD):
    try:
        with open(SNAPSHOT_OLD,encoding="utf-8") as f: snap_ant = json.load(f)
        ant_map = {(x["nombre"],x["telefono"]):x for x in snap_ant.get("contactos",[])}
        now_map = {(c["nombre"],c["telefono"]):c for c in contactos}
        fecha_ant = snap_ant.get("fecha","—")
        for key,c in now_map.items():
            if key not in ant_map:
                cambios.append({"tipo":"nuevo","desc":f"Nuevo contacto: <strong>{c['nombre']}</strong> ({c['estado']})"})
            else:
                a = ant_map[key]
                if a["estado"] != c["estado"]:
                    extra = f" — {c['plan']}" if c["estado"]=="Cerrado" and c["plan"] else ""
                    cambios.append({"tipo":"estado","desc":
                        f"<strong>{c['nombre']}</strong>: {a['estado']} → <strong>{c['estado']}</strong>{extra}"})
        for key,a in ant_map.items():
            if key not in now_map:
                cambios.append({"tipo":"eliminado","desc":f"Eliminado: {a['nombre']}"})
    except Exception as e:
        cambios = [{"tipo":"info","desc":f"Error leyendo snapshot: {e}"}]
else:
    cambios = [{"tipo":"info","desc":"Primera corrida — snapshot generado."}]
with open(SNAPSHOT_NEW,"w",encoding="utf-8") as f: json.dump(snap_nuevo,f,ensure_ascii=False,indent=2)
shutil.copy(SNAPSHOT_NEW, SNAPSHOT_OLD)

# ── KPI histórico ─────────────────────────────────────────────────────────────
kpi_hist = []
if os.path.exists(KPI_HISTORICO):
    try:
        with open(KPI_HISTORICO,encoding="utf-8") as f: kpi_hist = json.load(f)
    except: kpi_hist = []
_emb = {e: sum(1 for c in contactos if c["estado"]==e) for e in ESTADOS}
kpi_hist.append({"fecha":snap_nuevo["fecha"],"lead":_emb.get("Lead",0),
    "en_contacto":_emb.get("En Contacto",0),"cotizado":_emb.get("Cotizado",0),
    "cerrado":_emb.get("Cerrado",0),"descartado":_emb.get("Descartado",0),"total":len(contactos)})
with open(KPI_HISTORICO,"w",encoding="utf-8") as f: json.dump(kpi_hist,f,ensure_ascii=False,indent=2)

# ── agregados ─────────────────────────────────────────────────────────────────
total_c      = len(contactos)
embudo_tot   = {e: sum(1 for c in contactos if c["estado"]==e) for e in ESTADOS}
cerrados_tot = embudo_tot.get("Cerrado",0)
respondieron = embudo_tot.get("En Contacto",0)+embudo_tot.get("Cotizado",0)+cerrados_tot
desc_tot     = embudo_tot.get("Descartado",0)
tasa_cierre  = round(cerrados_tot/total_c*100,1) if total_c else 0
tasa_resp    = round(respondieron/total_c*100,1) if total_c else 0
conv_ef      = round(cerrados_tot/(cerrados_tot+desc_tot)*100,1) if (cerrados_tot+desc_tot) else 0
total_ingr   = sum(c["cuota"] for c in contactos if c["estado"]=="Cerrado")
ventas_com   = sum(1 for c in contactos if c["estado"]=="Cerrado" and c["tipo"].lower()!="plan empleado")
capitas_pe   = sum(c["capitas"] for c in contactos if c["estado"]=="Cerrado" and c["tipo"].lower()=="plan empleado")
total_gs100  = sum(1 for c in contactos if c["estado"]=="Cerrado" and c["plan"]=="GS100" and c["tipo"].lower()!="plan empleado")
total_gs200  = sum(1 for c in contactos if c["estado"]=="Cerrado" and c["plan"]=="GS200" and c["tipo"].lower()!="plan empleado")
pe_cuils = {c["cuil"] for c in contactos if c["estado"]=="Cerrado" and c["tipo"].lower()=="plan empleado"}
if miembros:
    miembros_com = [m for m in miembros if m["cuil_titular"] not in pe_cuils]
    capitas_tot  = len(miembros_com)
    ventas_miem  = sum(1 for m in miembros_com if m["rol"].lower() == "titular")
else:
    capitas_tot = sum(c["capitas"] for c in contactos if c["estado"]=="Cerrado" and c["tipo"].lower()!="plan empleado")
    ventas_miem = ventas_com
ratio_cap = round(capitas_tot / ventas_miem, 2) if ventas_miem else 0

MOTIVOS_OPCIONES = ["Preexistencia","Costos","Cambio de cobertura reciente","No le interesa el producto","No contacto"]
motivos_data = {m: sum(1 for c in contactos if c["estado"]=="Descartado" and c.get("motivo")==m) for m in MOTIVOS_OPCIONES}
ingr_por_venta  = round(total_ingr / ventas_com,  0) if ventas_com  else 0
ingr_por_capita = round(total_ingr / capitas_tot, 0) if capitas_tot else 0
meses_unicos = sorted(set(c["mes"] for c in contactos if c["mes"]),
    key=lambda m:(int(m.split("-")[1]),int(m.split("-")[0])))
ventas_mes = {}
for c in contactos:
    if c["estado"]!="Cerrado" or not c["mes"]: continue
    m = c["mes"]
    if m not in ventas_mes: ventas_mes[m]={"gs100":0,"gs200":0,"capitas":0,"ingresos":0}
    if c["plan"]=="GS100": ventas_mes[m]["gs100"]+=1
    if c["plan"]=="GS200": ventas_mes[m]["gs200"]+=1
    ventas_mes[m]["capitas"]+=c["capitas"]
    ventas_mes[m]["ingresos"]+=c["cuota"]
meses_sorted = sorted(ventas_mes.keys(), key=lambda m:(int(m.split("-")[1]),int(m.split("-")[0])))

# ── costos por periodo ────────────────────────────────────────────────────────
def calcular_costos_periodo(contactos_filtrados):
    ingr = sum(c["cuota"] for c in contactos_filtrados if c["estado"]=="Cerrado")
    asesoras = {}; total_com = 0
    for com in COMERCIALES:
        cc = [c for c in contactos_filtrados if c["comercial"]==com and c["estado"]=="Cerrado"]
        fc = sum(c["capitas"] for c in cc if c["canal"]=="TERRENO")
        dc = sum(c["capitas"] for c in cc if c["canal"]=="OSPIF")
        ft = get_tramo(fc, ESCALA_FUERA); dt = get_tramo(dc, ESCALA_DENTRO)
        cf = ft["comision"] if ft else 0; cd = dt["comision"] if dt else 0
        total_com += cf + cd
        asesoras[com] = {"fuera_cap":fc,"comision_fuera":cf,"dentro_cap":dc,"comision_dentro":cd,"total":cf+cd}
    descuento_gf = sum(
        max(0, c["capitas"] * PRECIO_BASE.get(c["plan"], 0) - c["cuota"])
        for c in contactos_filtrados
        if c["estado"]=="Cerrado" and c["tipo"].lower()!="plan empleado"
        and c["capitas"]>1 and c["cuota"]!=PRECIO_BASE.get(c["plan"],0)
    )
    return {"ingresos":ingr,"asesoras":asesoras,"total_comision":total_com,
            "resultado":ingr-total_com,"descuento_gf":descuento_gf}

costos_periodos = {"todos": calcular_costos_periodo(contactos)}
for m in meses_unicos:
    costos_periodos[m] = calcular_costos_periodo([c for c in contactos if c["mes"]==m])

asesoras_data = []; total_comisiones = 0
for com in COMERCIALES:
    cc=[c for c in contactos if c["comercial"]==com and c["estado"]=="Cerrado"]
    fc=sum(c["capitas"] for c in cc if c["canal"]=="TERRENO")
    dc=sum(c["capitas"] for c in cc if c["canal"]=="OSPIF")
    ft=get_tramo(fc,ESCALA_FUERA); dt=get_tramo(dc,ESCALA_DENTRO)
    com_t=(ft["comision"] if ft else 0)+(dt["comision"] if dt else 0)
    total_comisiones+=com_t
    asesoras_data.append({"nombre":com,"ventas":len(cc),"fuera_cap":fc,"dentro_cap":dc,
                          "ft":ft,"dt":dt,"comision":com_t,"ingresos":sum(c["cuota"] for c in cc)})

def badge_tramo(t):
    if not t: return '<span class="badge badge-grey">Sin tramo</span>'
    return f'<span class="badge badge-blue">{t["min"]}+ cáp. &middot; {ars(t["comision"])}</span>'

def cambio_icon(tipo):
    return {"nuevo":"🟢","estado":"🔄","eliminado":"🔴","info":"ℹ️"}.get(tipo,"•")

asesoras_rows = "".join(f"""<tr>
      <td><strong>{a["nombre"]}</strong></td>
      <td class="right">{a["ventas"]}</td>
      <td class="right">{a["fuera_cap"]}</td><td>{badge_tramo(a["ft"])}</td>
      <td class="right">{ars(a["ft"]["comision"]) if a["ft"] else "—"}</td>
      <td class="right">{a["dentro_cap"]}</td><td>{badge_tramo(a["dt"])}</td>
      <td class="right">{ars(a["dt"]["comision"]) if a["dt"] else "—"}</td>
      <td class="right bold-val">{ars(a["comision"])}</td>
      <td class="right">{ars(a["ingresos"])}</td>
    </tr>""" for a in asesoras_data)

ventas_mes_rows = "".join(
    f'<tr><td>{m}</td><td class="right">{ventas_mes[m]["gs100"]}</td>'
    f'<td class="right">{ventas_mes[m]["gs200"]}</td>'
    f'<td class="right">{ventas_mes[m]["gs100"]+ventas_mes[m]["gs200"]}</td>'
    f'<td class="right">{ventas_mes[m]["capitas"]}</td>'
    f'<td class="right">{ars(ventas_mes[m]["ingresos"])}</td></tr>'
    for m in meses_sorted
) or '<tr><td colspan="6" style="text-align:center;color:var(--muted)">Sin ventas cerradas aún</td></tr>'

cambios_rows = "\n".join(
    f'<tr class="changelog-row-{ch["tipo"]}"><td>{cambio_icon(ch["tipo"])}</td>'
    f'<td>{ch["desc"]}</td><td>{snap_nuevo["fecha"]}</td></tr>'
    for ch in cambios
) or '<tr><td colspan="3" style="text-align:center;color:var(--muted)">Sin cambios</td></tr>'

cambios_count = len(cambios)
period_btns = '<button class="period-btn active" onclick="setPeriodo(\'todos\',this)">Todos</button>' + \
    "".join(f'<button class="period-btn" onclick="setPeriodo(\'{m}\',this)">{m}</button>' for m in meses_unicos)
mes_options = '<option value="">Todos los meses</option>' + \
    "".join(f'<option value="{m}">{m}</option>' for m in meses_unicos)

js_contactos   = json.dumps(contactos, ensure_ascii=False)
js_costos      = json.dumps(costos_periodos, ensure_ascii=False)
js_meses       = json.dumps(meses_sorted)
js_gs100       = json.dumps([ventas_mes[m]["gs100"]    for m in meses_sorted])
js_gs200       = json.dumps([ventas_mes[m]["gs200"]    for m in meses_sorted])
js_ingresos_ch = json.dumps([ventas_mes[m]["ingresos"] for m in meses_sorted])
js_marketing   = json.dumps(marketing_data, ensure_ascii=False)
js_cobros_tabla= json.dumps(cobros_tabla_meses, ensure_ascii=False)
js_detalle_c   = json.dumps(detalle_cobros, ensure_ascii=False)
js_padron      = json.dumps(padron, ensure_ascii=False)
js_embudo_tot  = json.dumps([embudo_tot.get(e,0) for e in ESTADOS])
js_comerciales = json.dumps(COMERCIALES)
js_kpi_hist    = json.dumps(kpi_hist, ensure_ascii=False)
js_motivos     = json.dumps(motivos_data, ensure_ascii=False)

# ── NOTA: el bloque PHP/HTML se omite aquí para no duplicar.
# El archivo completo en producción está en:
# 2_TABLEROS/TABLERO COMERCIAL/2_scripts/actualizar.py
# La estructura es: php = f"""...""" con todo el HTML/JS embebido,
# luego: with open(OUT_PHP,"w",encoding="utf-8") as f: f.write(php)
# Leer ese archivo directamente para el bloque HTML completo.
print(f"  Contactos: {total_c} | Cerrados: {cerrados_tot} | Comisiones: {ars(total_comisiones)}")
```

> **Nota sobre el HTML/JS del tablero:** El bloque `php = f"""..."""` (líneas ~554-1493 del script original) genera el dashboard completo. Para un cliente nuevo, Claude puede regenerarlo desde cero con los mismos conceptos adaptados al negocio del cliente. Lo que sí es reutilizable 1:1 es toda la lógica Python de arriba (lectura de Excel, cálculos, costos, cobros).

---

*Fin del documento — con este archivo y la carpeta del proyecto, una sesión nueva de Claude puede entender, replicar y adaptar el sistema completo.*
