#!/usr/bin/env bash
# startup.sh

echo "===== Iniciando o startup script ====="

# 1) Instalar dependências Python
echo "Instalando pacotes do requirements.txt..."
pip install --upgrade pip
pip install --no-cache-dir -r requirements.txt

# 2) Instalar navegadores (Playwright)
# Se você só precisa do Chromium, pode usar: playwright install chromium
echo "Instalando navegadores do Playwright..."
playwright install

# 3) Subir a aplicação (uvicorn)
echo "Iniciando servidor com uvicorn..."
exec uvicorn app:app --host 0.0.0.0 --port 8000
