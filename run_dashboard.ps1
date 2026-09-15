$env:DASHFG_HEADLESS = 1
cd 'C:\Apps\Dash_ambiental_qualidade'
& py -3.12 gerar_dashboard_pbix.py 2>&1
