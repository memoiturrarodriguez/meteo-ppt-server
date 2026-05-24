# Servidor PPT Meteo Operacional - Deploy en Render

Este paquete permite dejar el generador de PowerPoint en la nube.
La app Flutter enviará el JSON al endpoint `/generar-ppt`, el servidor generará la PPT y la devolverá al teléfono para guardarla/compartirla.

## Archivos incluidos

- app.py: servidor Flask.
- generar_ppt_meteo_final_v3.py: generador PPT actual.
- requirements.txt: dependencias Python.
- render.yaml: configuración opcional para Render.
- .python-version: versión Python sugerida.

## Deploy recomendado: Render

1. Crea una cuenta en https://render.com
2. Crea un repositorio en GitHub, por ejemplo: meteo-ppt-server
3. Sube estos archivos al repositorio.
4. En Render: New > Web Service.
5. Conecta el repositorio GitHub.
6. Configura:
   - Language: Python 3
   - Build Command: pip install -r requirements.txt
   - Start Command: gunicorn app:app --timeout 180 --workers 1
7. Deploy.

Cuando Render termine, te dará una URL similar a:

https://meteo-ppt-server.onrender.com

La URL que debe usar la app será:

https://meteo-ppt-server.onrender.com/generar-ppt

## Prueba rápida

Abre en navegador:

https://TU-SERVICIO.onrender.com/

Debe responder algo como:
Servidor PPT Meteo funcionando correctamente.

## Nota importante

En planes gratuitos, Render puede dormir el servicio si no se usa.
La primera exportación puede demorar más por "arranque en frío".
