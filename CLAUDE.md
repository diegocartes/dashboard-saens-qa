# CLAUDE.md — Dashboard de Control de Protocolos QA/QC · TK-AB

Contexto y reglas para trabajar este proyecto con Claude Code. **Todo el proyecto y sus
textos están en español.**

## Qué es

Aplicación web que visualiza el avance de los protocolos de reparación del estanque
**TK-AB** (originalmente TK-005), proyecto *"Reparación TK de Hormigón – Tren AB"*.

- **Ejecutor:** SAENS Polímeros y Revestimientos Ltda.
- **Cliente:** Compañía Minera Zaldívar (CMZ · Antofagasta Minerals).
- **Contrato:** SC-1513 / SPO2514 · N° 4540008590 · Área 36 Tank Farm (SX/EW).
- **Revisor/Aprobador (ADC):** Andrés Arancibia.

La app se alimenta de la planilla `Mapa_Control_Protocolos_TK-AB.xlsx`: el usuario la
carga, la app lee la hoja **REGISTRO** y dibuja KPIs, el plano del estanque, la matriz de
avance y la distribución de estados. Los datos se guardan en `localStorage` y se recargan
al volver a abrir la página; al cargar otra planilla se reemplazan.

## Stack

- **Vite 5** + **React 18** + **Tailwind CSS 3**
- **xlsx (SheetJS)** para parsear el Excel en el navegador
- **lucide-react** para iconos
- **Express** (`server.js`) para servir el build en producción (Railway)

## Comandos

```bash
npm install        # instalar dependencias (requiere Node.js 18+)
npm run dev        # desarrollo (Vite, http://localhost:5173)
npm run build      # build de producción -> carpeta dist/
npm start          # sirve dist/ con Express en $PORT (o 3000)
npm run preview    # preview del build con Vite
```

## Estructura

```
index.html          charset UTF-8; monta /src/main.jsx
server.js           Express: static(dist) + fallback SPA; usa process.env.PORT
src/main.jsx        crea el root de React
src/index.css       directivas de Tailwind
src/Dashboard.jsx   TODO: parseo, cálculo y UI
```

`src/Dashboard.jsx` contiene:
- Constantes `C` (paleta), `EST` (estados+colores), `EST_ORDER`, `RANK`.
- Helpers puros: `parse()`, `numOrNull()`, `serialToStr()`, `protoMeta()`, `estadoOf()`,
  `heat()`, `pctTxt()`.
- Componentes **de nivel superior** que reciben props: `Ring`, `Seg`, `WallGrid`,
  `Matrix`, `EstadoBar`, `ProcBars`, `Inspector`, `Header`.
- `export default function Dashboard()`: único componente con estado/hooks.
- El logo SAENS va embebido como base64 en la constante `LOGO`.

## Modelo de datos (hoja REGISTRO)

- El encabezado se detecta buscando una celda que sea exactamente `Elemento` **y** otra que
  contenga `Avance` (así se evita falsear con la fila de título).
- Columnas usadas: `Elemento`, `Paño/Col`, `Nivel/Fila`, `Protocolo`, `% Avance`,
  `Ejecutó (SAENS)`, `ITO (CMZ)`, `Fecha Ctrl Interno`, `Result. Interno (PC)`,
  `Fecha Ctrl Final`, `Result. Final (PP)`, `N° NCR`, `Observaciones`.
- `% Avance` se guarda como fracción 0..1; si llega un valor > 1.5 se divide por 100.
- El **ESTADO se recalcula en JS** con `estadoOf()` (no se lee el valor cacheado del Excel).

### Estados (no cambiar las claves)
`FINAL` (Aprobado final PP) · `INTERNO` (Aprobado interno PC) · `C/OBS` (Aprob. c/obs) ·
`PROCESO` · `RECHAZADO` (NCR) · `NO INICIADO`.

## Dominio: protocolos y geometría

4 protocolos (gobernados por el PIE `1005-SPO2514-SC1513-06-36-QA-PIE-0002`):
1. Picado Muro y Piso — `1005-SPO2514-SC1513-06-36-QA-PRT-0001`
2. Instalación Enfierradura — `...-QA-PRT-0002`
3. Hormigón Polimérico — `...-QA-PRT-0003`
4. Aplicación FRP — `...-QA-PRT-0004`

Tipos de control: **PC** = Punto de Control (interno SAENS) · **PP** = Punto de Parada
(liberación final con CMZ/ITO).

Geometría de los sectores (la app la deduce de los datos, no la fija):
- **Muro Norte / Sur:** 3 niveles (N1–N3) × 10 paños (P01–P10).
- **Muro Oeste / Este:** 3 niveles × 4 paños (P01–P04).
- **Piso:** grilla columnas (C01–C10) × filas (F1–F4).

## Convenciones (importante)

- **Idioma:** español en UI, comentarios y entregables.
- **Nomenclatura:** preservar SIEMPRE los códigos exactos (PRT, PIE, contratos, IDs). No
  normalizar, traducir ni alterar.
- **UTF-8 real en JSX:** escribir los acentos y símbolos como caracteres reales
  (`Hormigón`, `paño`, `·`, `×`, `–`). **No** usar escapes `\uXXXX` dentro de texto JSX:
  esbuild/Vite los imprime literales (ese fue un bug ya corregido).
- **Componentes:** mantener todos los subcomponentes al nivel superior recibiendo props.
  No anidar definiciones de componentes dentro de otro componente.
- **`localStorage`:** clave `tkab_dashboard_v1`; siempre envuelto en try/catch.
- **Colores:** `heat()` blanco→verde para % de avance; colores de `EST` para estado.

## Despliegue en Railway

1. Subir el repo a GitHub.
2. Railway → New Project → Deploy from GitHub repo → seleccionar el repo.
3. Railway detecta Node y ejecuta `npm install` + `npm run build`, luego `npm start`.
   El puerto lo inyecta Railway vía `PORT` (ya contemplado en `server.js`).
4. No se requieren variables de entorno.
5. Settings → Networking → Generate Domain para obtener la URL pública.

La persistencia es por navegador/dispositivo (no es base de datos compartida): cada usuario
carga su planilla una vez y queda guardada localmente.
