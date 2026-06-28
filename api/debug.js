import { list } from '@vercel/blob';

export default async function handler(req, res) {
  const hasToken = !!process.env.BLOB_READ_WRITE_TOKEN;
  try {
    const result = await list({ prefix: 'dashboard-state.json' });
    return res.status(200).json({ hasToken, blobCount: result.blobs.length, blobs: result.blobs });
  } catch (e) {
    return res.status(200).json({ hasToken, error: e.message });
  }
}
