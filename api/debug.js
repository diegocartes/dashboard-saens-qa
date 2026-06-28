import { list } from '@vercel/blob';

export default async function handler(req, res) {
  // Buscar cualquier variable de entorno relacionada con blob/token
  const blobKeys = Object.keys(process.env).filter(k =>
    k.includes('BLOB') || k.includes('blob') || k.includes('TOKEN') || k.includes('VERCEL_')
  );

  const hasToken = !!process.env.BLOB_READ_WRITE_TOKEN;

  try {
    const result = await list({ prefix: 'dashboard-state.json' });
    return res.status(200).json({ hasToken, blobKeys, blobCount: result.blobs.length });
  } catch (e) {
    return res.status(200).json({ hasToken, blobKeys, error: e.message });
  }
}
