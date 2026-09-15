Set-Location 'C:\Apps\Dash_ambiental_qualidade'
$env:DASHFG_HEADLESS = 1
& 'py.exe' -3.12 'gerar_dashboard_pbix.py' 2>&1 | Out-Host
Write-Host "Geração concluída."
