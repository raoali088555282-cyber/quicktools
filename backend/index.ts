import { db } from '@appdeploy/sdk';
import { error, json, router } from '@appdeploy/sdk';

type QRLink = {
  target: string;
  createdAt: string;
};

type QRScan = {
  qrId: string;
  at: string;
  device: string;
};

export const handler = router({
  'POST /api/qr': [async ({ body }) => {
    const input = body as { target?: string };
    if (!input?.target) return error('Destination URL is required', 400);
    let target: URL;
    try {
      target = new URL(input.target);
    } catch {
      return error('Invalid destination URL', 400);
    }
    if (!/^https?:$/.test(target.protocol)) return error('Only HTTP and HTTPS URLs are supported', 400);
    const [id] = await db.add<QRLink>('qr_links', [{ target: target.toString(), createdAt: new Date().toISOString() }]);
    if (!id) return error('Could not create QR link', 500);
    return json({ id, target: target.toString() });
  }],
  'GET /api/qr/:id/analytics': [async ({ params }) => {
    const result = await db.list<QRScan>('qr_scans', { limit: 100, filter: { qrId: params.id } });
    const scans = result.items
      .map(item => ({ at: item.at, device: item.device }))
      .sort((a, b) => b.at.localeCompare(a.at));
    return json({ scans, total: scans.length });
  }],
  'GET /api/qr/:id/redirect': [async ({ params, event }) => {
    const [record] = await db.get<QRLink>('qr_links', [params.id]);
    if (!record?.target) return error('QR link not found', 404);
    const headers = event?.headers || {};
    const userAgent = String(headers['user-agent'] || headers['User-Agent'] || '');
    const device = /Mobi|Android|iPhone|iPad/i.test(userAgent) ? 'mobile' : 'desktop';
    await db.add<QRScan>('qr_scans', [{ qrId: params.id, at: new Date().toISOString(), device }]);
    return {
      statusCode: 302,
      headers: { Location: record.target, 'Cache-Control': 'no-store' },
      body: ''
    };
  }]
});
