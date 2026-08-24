#!/usr/bin/env bash
# ============================================================
#  ACTUALIZAR Y PUBLICAR - Griff Salud
# ============================================================
#  Corre todos los tableros, copia los outputs a las carpetas
#  publicas y hace push a GitHub.
# ============================================================
set -e
cd "$(dirname "$0")"

echo "============================================================"
echo "  1) ACTUALIZAR TABLEROS"
echo "============================================================"

# Cuentas a Pagar
echo "[1/5] Cuentas a Pagar..."
( cd "2_TABLEROS/TABLERO CUENTAS A PAGAR CLOUD" && python 2_scripts/actualizar.py )

# Estado de Resultado
echo "[2/5] Estado de Resultado..."
( cd "2_TABLEROS/TABLERO ESTADO DE RESULTADO" && python 2_scripts/actualizar.py )

# Facturacion
echo "[3/5] Facturacion..."
( cd "2_TABLEROS/TABLERO FACTURACION" && python actualizar.py )

# Flujo Bancario
echo "[4/5] Flujo Bancario..."
( cd "2_TABLEROS/TABLERO FLUJO BANCARIO" && python 2_scripts/actualizar.py )

# [OCULTO] Proyeccion de Pagos — desactivado, archivos conservados
# ( cd "2_TABLEROS/TABLERO PROYECCION DE PAGOS" && python 2_scripts/actualizar.py )

echo "============================================================"
echo "  2) PUBLICAR (copiar outputs a carpetas publicas)"
echo "============================================================"

# Helper: convertir HTML/PHP de origen a PHP de destino agregando auth_check
publish() {
  local SRC_FILE="$1" DEST_FILE="$2"
  if [ -f "$SRC_FILE" ]; then
    echo "  ${SRC_FILE##*/} -> $DEST_FILE"
    {
      echo '<?php require_once __DIR__ . "/../auth_check.php"; ?>'
      # Si el archivo ya tiene <?php require ... auth_check ... lo quitamos del inicio
      sed -e '1{/<?php.*auth_check.*?>/d;}' "$SRC_FILE"
    } > "$DEST_FILE"
  fi
}

# Cuentas a Pagar
mkdir -p cuentas_a_pagar
cp -f "2_TABLEROS/TABLERO CUENTAS A PAGAR CLOUD/3_outputs"/*.csv cuentas_a_pagar/ 2>/dev/null || true
cp -f "2_TABLEROS/TABLERO CUENTAS A PAGAR CLOUD/3_outputs"/*.xlsx cuentas_a_pagar/ 2>/dev/null || true
cp -f "2_TABLEROS/TABLERO CUENTAS A PAGAR CLOUD/3_outputs"/log_ultima_corrida.txt cuentas_a_pagar/ 2>/dev/null || true
publish "2_TABLEROS/TABLERO CUENTAS A PAGAR CLOUD/3_outputs/tablero.php" cuentas_a_pagar/tablero.php

# Estado de Resultado
mkdir -p estado_resultado
cp -f "2_TABLEROS/TABLERO ESTADO DE RESULTADO/3_outputs"/*.xlsx estado_resultado/ 2>/dev/null || true
cp -f "2_TABLEROS/TABLERO ESTADO DE RESULTADO/3_outputs"/log_ultima_corrida.txt estado_resultado/ 2>/dev/null || true
publish "2_TABLEROS/TABLERO ESTADO DE RESULTADO/3_outputs/dashboard.php" estado_resultado/dashboard.php

# Facturacion
mkdir -p facturacion
cp -f "2_TABLEROS/TABLERO FACTURACION/datos.js" facturacion/ 2>/dev/null || true
publish "2_TABLEROS/TABLERO FACTURACION/dashboard.php" facturacion/dashboard.php

# Flujo Bancario
mkdir -p flujo_bancario
cp -f "2_TABLEROS/TABLERO FLUJO BANCARIO/3_outputs"/*.csv flujo_bancario/ 2>/dev/null || true
cp -f "2_TABLEROS/TABLERO FLUJO BANCARIO/3_outputs"/*.xlsx flujo_bancario/ 2>/dev/null || true
cp -f "2_TABLEROS/TABLERO FLUJO BANCARIO/3_outputs"/log_ultima_corrida.txt flujo_bancario/ 2>/dev/null || true
publish "2_TABLEROS/TABLERO FLUJO BANCARIO/3_outputs/tablero.php" flujo_bancario/tablero.php

# Proyeccion de Pagos
mkdir -p proyeccion
cp -f "2_TABLEROS/TABLERO PROYECCION DE PAGOS/2_scripts"/*.xlsx proyeccion/ 2>/dev/null || true
publish "2_TABLEROS/TABLERO PROYECCION DE PAGOS/2_scripts/plantilla.php" proyeccion/proyeccion.php

echo "============================================================"
echo "  3) PUBLICAR EN GITHUB"
echo "============================================================"

git add -A
git -c user.email="tablero@griffsalud.com.ar" -c user.name="Tablero Griff" \
    commit -m "Actualizacion automatica $(date +'%Y-%m-%d %H:%M')" || echo "  (sin cambios para commitear)"
git push origin main || git push origin master || echo "  ATENCION: no pude pushear. Revisa el remoto."

echo "============================================================"
echo "  LISTO"
echo "============================================================"
