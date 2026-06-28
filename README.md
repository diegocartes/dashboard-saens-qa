# Dashboard de Control de Protocolos QA/QC · TK-AB

Dashboard web para visualizar el avance de los protocolos de reparación del estanque
**TK-AB** (Reparación TK de Hormigón – Tren AB · SC-1513 / SPO2514 · SAENS / CMZ).

Se alimenta de la planilla `Mapa_Control_Protocolos_TK-AB.xlsx`: el usuario la carga,
la app lee la hoja **REGISTRO** y dibuja el plano del estanque, KPIs, matriz y estados.
Los datos quedan guardados en el navegador (`localStorage`) y se recargan solos al volver
a abrir la página. Al cargar otra planilla se reemplazan.

## Requisitos

- Node.js 18 o superior

## Desarrollo local

```bash
npm install
npm run dev
```

Abre la URL que muestra Vite (por defecto http://localhost:5173).

## Build de producción

```bash
npm run build   # genera la carpeta dist/
npm start       # sirve dist/ con Express en el puerto $PORT (o 3000)
```

## Publicar en GitHub

```bash
git init
git add .
git commit -m "Dashboard TK-AB"
git branch -M main
git remote add origin https://github.com/<tu-usuario>/tk-ab-dashboard.git
git push -u origin main
```

## Desplegar en Railway

1. Entra a [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**.
2. Selecciona el repositorio `tk-ab-dashboard`.
3. Railway detecta Node automáticamente y ejecuta:
   - **Build:** `npm install` y `npm run build`
   - **Start:** `npm start` (lee el puerto desde la variable `PORT` que inyecta Railway).
4. No se necesitan variables de entorno adicionales.
5. En **Settings → Networking → Generate Domain** se obtiene la URL pública.

## Notas

- La persistencia es por navegador/dispositivo (no es una base de datos compartida).
  Cada usuario carga su planilla una vez y la mantiene localmente.
- El logo SAENS va embebido en el componente (`src/Dashboard.jsx`).
- Toda la lógica de parseo y cálculo está en `src/Dashboard.jsx`.
