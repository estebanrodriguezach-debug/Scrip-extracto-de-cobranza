#!/usr/bin/env bash
# Deja todo listo para correr el volcado desde Google Cloud Shell.
#
#   bash scripts/setup_cloudshell.sh
#
# Cloud Shell ya viene autenticada con tu usuario, asi que NO hace falta crear
# una cuenta de servicio ni descargar ningun JSON de clave.

set -euo pipefail

PROYECTO="${PROYECTO:-tranquil-wave-503300-d6}"  # ACH FACTURACION
ALCANCES="https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/drive.file,https://www.googleapis.com/auth/cloud-platform"

echo "==> Proyecto: ${PROYECTO}"
gcloud config set project "${PROYECTO}"

echo "==> Habilitando Google Sheets API y Google Drive API"
gcloud services enable sheets.googleapis.com drive.googleapis.com

echo "==> Entorno de Python"
python3 -m venv .venv
# shellcheck disable=SC1091
source .venv/bin/activate
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

echo
echo "==> Falta un paso manual: autorizar el acceso a tus planillas."
echo "    Corre esto y segui el link que te muestra:"
echo
echo "      gcloud auth application-default login --scopes=${ALCANCES}"
echo
echo "    Despues:"
echo "      source .venv/bin/activate"
echo "      python3 -m clasificador.cli crear --titulo 'Extractos de cobranzas'"
echo "      python3 -m clasificador.cli volcar Extractos_cobranzas.xlsx --planilla <ID>"
