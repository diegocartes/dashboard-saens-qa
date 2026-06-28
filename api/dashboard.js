import { put, list, del } from '@vercel/blob';

const BLOB = 'dashboard-state.json';
const EMPTY = { recs: [], fileName: '', savedAt: '' };

export default async function handler(req, res) {
  if (req.method === 'GET') {
    try {
      const { blobs } = await list({ prefix: BLOB });
      if (!blobs.length) return res.status(200).json(EMPTY);
      const data = await fetch(blobs[0].url, {
        headers: { Authorization: `Bearer ${process.env.BLOB_READ_WRITE_TOKEN}` }
      }).then(r => r.json());
      return res.status(200).json(data);
    } catch { return res.status(200).json(EMPTY); }
  }

  if (req.method === 'POST') {
    try {
      const body = req.body;
      console.log('POST body type:', typeof body, 'keys:', body ? Object.keys(body).join(',') : 'null');
      const content = JSON.stringify(body);
      console.log('POST content length:', content?.length);
      await put(BLOB, content, { addRandomSuffix: false });
      return res.status(200).json({ ok: true });
    } catch (e) {
      console.error('POST error:', e.message);
      return res.status(500).json({ error: e.message });
    }
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
