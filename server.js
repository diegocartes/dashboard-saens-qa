import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const dist = path.join(__dirname, 'dist');

app.use(express.static(dist));
// SPA fallback
app.get('*', (_req, res) => res.sendFile(path.join(dist, 'index.html')));

const port = process.env.PORT || 3000;
app.listen(port, '0.0.0.0', () => console.log('Dashboard TK-AB escuchando en puerto ' + port));
