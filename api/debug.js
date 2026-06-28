import { put, list } from '@vercel/blob';

export default async function handler(req, res) {
  const hasToken = !!process.env.BLOB_READ_WRITE_TOKEN;

  // Test list
  let listResult = null, listError = null;
  try {
    const r = await list({ prefix: 'dashboard-state.json' });
    listResult = r.blobs.length;
  } catch (e) { listError = e.message; }

  // Test put con payload mínimo
  let putOk = false, putError = null;
  try {
    await put('debug-test.json', JSON.stringify({ test: true }), { access: 'public', addRandomSuffix: false });
    putOk = true;
  } catch (e) { putError = e.message; }

  return res.status(200).json({ hasToken, listResult, listError, putOk, putError });
}
