# Atajos Globales

Panel Angular para ejecutar manualmente los bots de Galleria, Posco y el flujo de inventario de cajas de Kometsales.

La aplicación usa un servidor Node/Express para que el token de GitHub nunca llegue al navegador. El mismo servidor entrega la compilación de Angular, por lo que se despliega como un solo Web Service en Render.

## Bot Pegar Data

El prototipo se encuentra en `bot/pegar_data.py`. En su primera etapa:

1. Abre Posco e inicia sesión.
2. Abre el menú **Órdenes**.
3. Selecciona la opción **Órdenes**.
4. Guarda capturas de cada paso como artifact de GitHub Actions durante 7 días.

El workflow manual está en `.github/workflows/pegar-data.yml` y requiere estos secretos en `BOT-123`:

- `POSCO_USER`
- `POSCO_PASSWORD`

Las credenciales de SharePoint se agregarán cuando se implemente la etapa de actualización del archivo `.xlsm`; nunca deben guardarse en archivos del repositorio.

## Bot Subir XLS Cajas

El botón **Subir XLS Cajas** descarga el archivo `Inventory Upload Boxes 050926.xlsx` desde SharePoint, normaliza el formato de `Available From`, entra a Kometsales, selecciona `PACIFICA FARMS - CAL`, vacía el inventario de cajas y sube el XLS.

El workflow está en `.github/workflows/inventory-boxes.yml` y requiere estos secretos en `BOT-123`:

- `KOMET_USER`
- `KOMET_PASSWORD`
- `SHAREPOINT_TENANT_ID`
- `SHAREPOINT_CLIENT_ID`
- `SHAREPOINT_CLIENT_SECRET`

El botón solicita `CONFIRMAR` antes de enviarlo a GitHub Actions porque el proceso elimina el inventario existente.

## Configuración local

Requisitos: Node.js 22.12 o posterior y npm.

1. Copia `.env.example` como `.env`.
2. Configura `GITHUB_TOKEN` con un token fine-grained que tenga acceso a `BOT-GALLERIA` y `BOT-123`, con permisos **Actions: Read and write** y **Contents: Read-only**.
3. Cambia `APP_PASSWORD` por una contraseña privada para proteger los botones.
4. Instala y ejecuta:

```powershell
npm ci
npm run build
npm start
```

Abre `http://localhost:3000`.

Para trabajar sólo en la interfaz con recarga automática usa `npm run dev`. Las llamadas `/api` requieren que el servidor Node esté disponible; la prueba completa siempre debe hacerse con la compilación y `npm start`.

## Despliegue en Render

El archivo `render.yaml` contiene toda la configuración del servicio.

1. Sube este proyecto a un repositorio nuevo de GitHub. No incluyas ningún archivo `.env`.
2. En Render selecciona **New > Blueprint** y conecta ese repositorio.
3. Render solicitará los valores secretos:
   - `GITHUB_TOKEN`: token nuevo de GitHub con los permisos indicados arriba.
   - `APP_PASSWORD`: contraseña que usarás para entrar al panel.
4. Crea el servicio y espera a que el health check `/api/health` sea correcto.

No reutilices el token que estaba escrito en la URL remota de `BOT-GALLERIA`: revócalo y crea uno nuevo.

## Comandos de verificación

```powershell
npm run build
npm test -- --watch=false
npm audit --omit=dev
```
