# Script de inicio para PowerShell
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SISTEMA DE EXTRACCION DE PAGOS WHATSAPP - EXCEL" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Activando entorno virtual..." -ForegroundColor Yellow
.\venv\Scripts\Activate.ps1
Write-Host ""
Write-Host "Iniciando sistema..." -ForegroundColor Green
python main.py

