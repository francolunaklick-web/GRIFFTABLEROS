# Contexto: Tablero Control de Aportes — Griff Salud
> Documento para cargar en una nueva sesión de Claude junto con la carpeta del proyecto.
> Última actualización: octubre 2026

---

## 1. QUÉ ES ESTE TABLERO

El **Control de Aportes** es un dashboard web que procesa los aportes y contribuciones de empleadores a OSPIF (Obra Social), cruzando dos fuentes principales:
- **Archivo de aportes**: qué personas aportaron, cuánto y en qué período
- **Padrón de afiliados**: a qué grupo familiar pertenece cada persona

El resultado es una vista por **Grupo Familiar**, con per cápita, distribución por delegación, evolución mensual, y detección de personas que aportaron pero no están identificadas en el padrón.

**El sistema NO tiene base de datos.** Los datos viven en archivos Excel. Un script Python los procesa y genera un archivo `datos.js` que el dashboard web carga dinámicamente.

---

## 2. ARCHIVOS PRINCIPALES

### Carpeta raíz del proyecto
```
TABLERO GENERAL - GRIFF WEB/
├── aportes/
│   ├── dashboard.php          ← Dashboard web (HTML/CSS/JS, lee datos.js)
│   └── datos.js               ← Generado por actualizar.py. NO editar a mano.
├── 1_DATOS/aportes/           ← Archivos de entrada (Excel)
│   ├── aportes_afiliados.xlsx    ← Aportes del período (cambia con cada actualización)
│   ├── afiliados_padron.xlsx     ← Padrón activo (cambia con cada actualización)
│   ├── padron_historico.xlsx     ← Padrón histórico de OSPIF (nombre fijo, no rota)
│   └── bajas_padron.xlsx         ← Padrón completo de bajas (nombre fijo, no rota)
└── 2_TABLEROS/TABLERO APORTES/2_scripts/
    └── actualizar.py          ← Script principal. Franco lo corre desde VS Code.
```

### CRÍTICO: el usuario (Franco) corre `actualizar.py` manualmente desde Visual Studio Code.
**Claude NUNCA debe correr `ACTUALIZAR_TODO.bat` ni el script principal del tablero.**

---

## 3. FLUJO DE DATOS

```
1_DATOS/aportes/
├── aportes_afiliados.xlsx   ─┐
├── afiliados_padron.xlsx    ─┤→ actualizar.py → aportes/datos.js → aportes/dashboard.php
├── padron_historico.xlsx    ─┤
└── bajas_padron.xlsx        ─┘
```

Cuando hay datos nuevos:
1. Franco reemplaza `aportes_afiliados.xlsx` y `afiliados_padron.xlsx` en `1_DATOS/aportes/`
2. Corre `actualizar.py` desde VS Code
3. El script regenera `datos.js`
4. Hace `git push` → Hostinger actualiza la web automáticamente

---

## 4. ESTRUCTURA DE LOS ARCHIVOS EXCEL

Todos los Excel tienen **encabezados en la fila 5 (índice 4 en Python)**, es decir, las primeras 4 filas son título/fecha/blancos/categoría. Siempre leer con `header=4`.

### `aportes_afiliados.xlsx` (varía en nombre, se busca el más reciente con glob `aportes*.xlsx`)
Columnas relevantes:
- `CUIL Empleado` — CUIL de la persona que aportó (dtype str para preservar ceros)
- `Nombre Empleado` — nombre
- `Delegación` — delegación OSPIF
- `Convenio` — convenio (todos son OSPIF, incluyendo CCENT, SAL, etc.)
- `Importe Aportes` — monto de aportes del empleado
- `Importe Contribuciones` — monto de contribuciones del empleador
- `Importe Total` — total (puede diferir de Aportes + Contribuciones si hay otros conceptos)
- `Fecha Lote` — fecha del lote (se usa para extraer el mes)

### `afiliados_padron.xlsx` (varía en nombre, se busca el más reciente con glob `afiliados*.xlsx`)
Columnas relevantes:
- `Credencial` — clave del grupo familiar. Formato: primeros N-2 dígitos = Grupo ID, últimos 2 = número de miembro (00 = titular)
- `Nº Documento` — DNI (dtype str)
- `Nombre` — nombre del afiliado
- `Delegación` — delegación
- `Parentesco` — relación con el titular

### `padron_historico.xlsx` (nombre fijo)
Columnas relevantes:
- `Cuil Titular` — CUIL del titular (se extrae DNI: posiciones [2:10] sin ceros iniciales)
- `Estado` — Baja / Activo / Desafiliado / Bloqueo
- `Fecha Estado` — fecha del cambio de estado
- `Apellidos y Nombres` — nombre
- `Delegación Anterior` — delegación

### `bajas_padron.xlsx` (nombre fijo, es el archivo BAJAS.xlsx grande — 212,947 filas)
Columnas relevantes:
- `Nº Documento` — DNI directo (dtype str, lstrip '0')
- `Cuil` — CUIL alternativo si no hay Nº Documento
- `Nombre` — nombre
- `Fecha Baja` — fecha de la baja
- `Delegación` — delegación
- `Convenio` — convenio

---

## 5. LÓGICA DE PROCESAMIENTO (actualizar.py)

### 5.1 Extracción de DNI
- De CUIL: `cuil[2:10].lstrip('0')` (elimina código de país + verificador, quita ceros iniciales)
- De Nº Documento: `str.strip().lstrip('0')`

### 5.2 Grupo Familiar
La `Credencial` determina el grupo:
- Credencial zerofill a 6 dígitos
- `Grupo ID` = primeros N-2 caracteres
- `Miembro Num` = últimos 2 caracteres (`00` = titular, `01`+ = a cargo)

### 5.3 Per cápita
```
per_capita = suma_importe_total_del_grupo / cantidad_miembros_del_grupo (del padrón)
```
Los miembros se cuentan desde el **padrón activo**, no desde los aportantes.

### 5.4 Rangos
- `menos25`: per cápita < $25,000
- `entre25y50`: $25,000 ≤ per cápita ≤ $50,000
- `mas50`: per cápita > $50,000

### 5.5 KPI Miembros (promedio mensual)
Cuando hay múltiples meses seleccionados, el KPI "Total Miembros" muestra el promedio mensual (total miembros / cantidad de meses), no la suma total.

---

## 6. CLASIFICACIÓN DE PERSONAS "SIN PADRÓN"

Las personas que aparecen en aportes pero **no se encuentran en el padrón activo** pasan por 3 capas de cruce:

```
Persona en aportes
    │
    ├─ ¿Está en afiliados_padron.xlsx? → SÍ → Se asigna a su grupo familiar (flujo normal)
    │
    └─ NO → "Sin padrón" → 3 capas de clasificación:
        │
        ├─ Capa 1: padron_historico.xlsx
        │   ├─ Estado "Baja" con fecha → "Baja 2026" o "Baja anterior (AAAA)"
        │   ├─ Estado "Desafiliado" → "Desafiliado"
        │   └─ Estado "Bloqueo" → "Bloqueo"
        │
        ├─ Capa 2: ¿El mes máximo del archivo coincide con algún mes de la persona?
        │   └─ SÍ → "Activo (sin padrón)" [aparece en aportes actuales pero sin registro]
        │
        ├─ Capa 3: bajas_padron.xlsx (212k filas)
        │   └─ Si DNI está aquí → "Baja 2026" o "Baja anterior (AAAA)" según año de Fecha Baja
        │
        └─ No encontrado en ninguno → "Sin clasificar"
```

**Resultado esperado con datos actuales:**
- Baja 2026 / Baja anterior: ~2,496 DNIs (encontrados en histórico o BAJAS)
- Activo (sin padrón): algunos
- Sin clasificar: ~2,486 — casi todos de CCENT (probable convenio especial sin afiliados en OSPIF)

**CCENT no es una clasificación de error.** Es una delegación/convenio interno de OSPIF. Los registros CCENT que quedan en "Sin clasificar" son personas que aportaron pero no tienen registro en ningún padrón conocido. Puede ser por convenio especial o datos no cargados.

---

## 7. OUTPUT: datos.js

El script genera un archivo JS con 3 variables globales:

```javascript
// Generado por actualizar.py — DD/MM/AAAA HH:MM
const APORTES_META = {
  actualizado: "...",
  arch_aportes: "...",
  arch_afil: "...",
  meses: ["07-2026", "08-2026", "09-2026"],
  delegaciones: ["OSPIF BAHIA BLANCA", ...],
  excluir_def: ["OSPIF ROSARIO", "OSPIF MENDOZA"],  // excluidos por defecto al cargar
  total_grupos: 1234,
  total_reg: 3456,
  sin_grupo: 4982,
  sin_dni: 4982,
  mes_max: "09-2026",
  sin_estados: { "Baja 2026": 120, "Baja anterior (2024)": 80, ... }
};

const APORTES_DATA = [
  // Un registro por grupo por mes
  {
    gid: "18760",          // Grupo ID
    titular: "APELLIDO, NOMBRE",
    deleg: "OSPIF BAHIA BLANCA",
    mes: "08-2026",
    tot: 39582.00,         // total del grupo ese mes
    ap: 0,                 // aportes (puede ser 0 si solo hay contribuciones u otros)
    co: 0,                 // contribuciones
    miembros: 1,           // del padrón activo
    a_cargo: 0,
    n_ap: 1,               // aportantes únicos ese mes
    pc: 39582.00,          // per cápita
    rango: "entre25y50",
    paren: ""              // parentescos de los a cargo
  }, ...
];

const APORTES_SIN_GRUPO = [
  // Un registro por DNI único (agrupado, no por registro de aporte)
  {
    cuil: "27251345630",
    dni: "25134563",
    nombre: "APELLIDO, NOMBRE",
    deleg: "OSPIF BAHIA BLANCA",
    convenio: "O.S.P.I.F.",
    estado: "Baja 2026",       // o "Baja anterior (2024)", "Activo (sin padrón)", "Sin clasificar"
    fecha_baja: "15/03/2026",  // vacío si no aplica
    meses_ap: "07-2026, 08-2026",
    n_meses: 2,
    tot: 19791.10,
    ap: 0,
    co: 0,
    con_ap: false              // true si ap > 0
  }, ...
];
```

---

## 8. DASHBOARD (dashboard.php)

### Pestañas
| Pestaña | Qué muestra | Fuente |
|---|---|---|
| Por Delegación | KPIs + tabla resumen agrupada por delegación | `APORTES_DATA` |
| Grupos Familiares | Tabla detallada por grupo con filtros internos | `APORTES_DATA` |
| Por Mes | Evolución mensual (grupos, miembros, totales) | `APORTES_DATA` |
| 📁 Bajas | Personas con estado Baja/Desafiliado/Activo-sin-padrón | `APORTES_SIN_GRUPO` filtrado |
| ❓ Sin clasificar | Personas sin registro en ningún padrón | `APORTES_SIN_GRUPO` filtrado |

### Filtros globales (afectan todas las pestañas)
- **Período (mes)**: selección múltiple de meses
- **Delegaciones**: multi-select con checkbox y búsqueda (vacío = todas)
- **Per cápita**: rango desde/hasta + botones de atajo (<25k, 25-50k, >50k)
- **Tamaño grupo**: mínimo de miembros

### Filtros locales (tab Grupos Familiares)
- Mes, delegación, rango de miembros, per cápita, orden asc/desc

### Descarga Excel
- Tab Grupos: exporta Grupos + Resumen por Delegación (2 hojas)
- Tab Bajas: exporta filtrado de bajas
- Tab Sin clasificar: exporta filtrado de sin clasificar
- Usa la librería SheetJS (XLSX) cargada desde CDN

### Estado JS global
```javascript
let F = { meses: [], delegs: [], pcMin: null, pcMax: null, tam: 0 };
// delegs vacío = todas las delegaciones incluidas
let GF = { mes:'', deleg:'', tam:0, pcMin:null, pcMax:null, orden:'asc' };
// GF = filtros locales del tab Grupos Familiares
```

---

## 9. CONVENCIONES IMPORTANTES

- **CCENT, SAL y similares son delegaciones internas de OSPIF**, no convenios externos. No clasificar diferente.
- Un registro con `Importe Aportes = $0` e `Importe Contribuciones = $0` pero `Importe Total > $0` es válido — puede haber otros conceptos.
- Los rangos per cápita se calculan sobre el **total del grupo**, no sobre aportes solamente.
- El per cápita divide por **miembros del padrón activo**, no por aportantes. Un grupo de 4 miembros paga per cápita = total / 4 aunque solo 1 aporte.

---

## 10. CÓMO PUBLICAR EN HOSTINGER

### Primera vez (setup)
1. Crear hosting en Hostinger (plan Business o superior — necesita PHP y Git)
2. En Panel de Hostinger → **Git** → conectar repositorio GitHub
   - Repo: `https://github.com/francolunaklick-web/GRIFFTABLEROS`
   - Branch: `main`
   - Directorio: raíz del dominio (`public_html/`)
3. En el File Manager de Hostinger, crear `config/clave.php` con el hash de la contraseña:
   ```bash
   # Generar hash desde terminal Hostinger o SSH:
   php -r "echo password_hash('la_clave_elegida', PASSWORD_DEFAULT);"
   ```
   Pegar en:
   ```php
   <?php
   $CLAVE_HASH = 'el_hash_generado';
   ```
4. Verificar que `config/.htaccess` exista con `Deny from all`

### Actualización rutinaria (después de correr actualizar.py)
```bash
git add -A
git commit -m "Aportes actualizado MM-AAAA"
git push
```
Hostinger detecta el push y actualiza la web en segundos. El tablero queda visible en el dominio configurado.

### Qué NO sube a GitHub (ver .gitignore)
- `config/clave.php` — contraseña
- `1_DATOS/**/*.xlsx` — Excel con datos del cliente
- `__pycache__/`, `*.pyc`
- Archivos de bloqueo (`.~lock.*`)

---

## 11. TROUBLESHOOTING FRECUENTE

| Síntoma | Causa probable | Solución |
|---|---|---|
| `KeyError: 'Columna X'` en actualizar.py | El Excel tiene `header` diferente o la columna cambió de nombre | Leer con `print(df.columns)` para ver las columnas reales |
| `datos.js` no cambia al actualizar | El script falló silenciosamente | Revisar la salida de la terminal de VS Code |
| Web no actualiza tras git push | Hostinger no se conectó bien | Ir al Panel Hostinger → Git → Deploy manual |
| "Sin clasificar" con muchos registros CCENT | Normal — CCENT no tiene afiliados en ningún padrón | Investigar si CCENT tiene su propio archivo de padrón |
| Grupos con per cápita = 0 | El grupo está en padrón pero sin aportes ese mes | Normal, se puede filtrar |
| Aportes = $0 pero Total > $0 | El importe está en otro concepto (diferente de aportes/contribuciones) | Es dato válido del archivo fuente |

---

## 12. DEPENDENCIAS PYTHON

```bash
pip install pandas openpyxl --break-system-packages
```

El script `actualizar.py` usa exclusivamente `pandas` y `openpyxl` (ya incluido en pandas).

---

## 13. ACCESO WEB

- **URL local (desarrollo):** No aplica — el dashboard es estático (datos embebidos en JS)
- **URL productiva:** a través del dominio configurado en Hostinger
- **Login:** contraseña única, gestionada en `config/clave.php`
- **Auth:** `auth_check.php` se incluye al inicio de cada `.php`. Si no hay sesión, redirige a `login.php`
