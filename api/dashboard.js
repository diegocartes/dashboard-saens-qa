import { put, list, del } from '@vercel/blob';

const BLOB = 'dashboard-state.json';
const EMPTY = { recs: [], fileName: '', savedAt: '' };

export default async function handler(req, res) {
  if (req.method === 'GET') {
    try {
      const { blobs } = await list({ prefix: BLOB });
      if (!blobs.length) return res.status(200).json(EMPTY);
      const data = await fetch(blobs[0].url).then(r => r.json());
      return res.status(200).json(data);
    } catch { return res.status(200).json(EMPTY); }
  }

  if (req.method === 'POST') {
    try {
      await put(BLOB, JSON.stringify(req.body), { access: 'public', addRandomSuffix: false });
      return res.status(200).json({ ok: true });
    } catch (e) { return res.status(500).json({ error: e.message }); }
  }

  if (req.method === 'DELETE') {
    try {
      const { blobs } = await list({ prefix: BLOB });
      for (const b of blobs) await del(b.url);
      return res.status(200).json({ ok: true });
    } catch (e) { return res.status(500).json({ error: e.message }); }
  }

  res.status(405).end();
}
