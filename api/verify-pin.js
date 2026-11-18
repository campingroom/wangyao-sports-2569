const crypto = require('crypto');

const DEFAULT_ADMIN_SALT = '67816bd3-d9b5-41ac-800f-a45f47154ec9';
const DEFAULT_ADMIN_PIN_HASH = '3368b75079a3745062e4d6e5286e0a271fd9a17fa8acf9c856a096a9cd1b39a2';
const MAX_FAILS = 5;
const LOCK_WINDOW_MS = 60 * 1000;
const attemptsByIp = new Map();

function getClientIp(req) {
  const xff = req.headers && (req.headers['x-forwarded-for'] || req.headers['x-real-ip']);
  if (Array.isArray(xff)) return String(xff[0] || 'unknown');
  if (typeof xff === 'string' && xff.trim()) return xff.split(',')[0].trim();
  return 'unknown';
}

function pruneAttempts(now) {
  if (attemptsByIp.size < 256) return;
  for (const [ip, rec] of attemptsByIp.entries()) {
    if (now - rec.firstAt > LOCK_WINDOW_MS) attemptsByIp.delete(ip);
  }
}

async function readJsonBody(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  if (typeof req.body === 'string') {
    try { return JSON.parse(req.body); } catch (_) { return {}; }
  }
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  if (!chunks.length) return {};
  try { return JSON.parse(Buffer.concat(chunks).toString('utf8')); } catch (_) { return {}; }
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, message: 'Method Not Allowed' });
  }

  const now = Date.now();
  pruneAttempts(now);
  const ip = getClientIp(req);
  const entry = attemptsByIp.get(ip);
  if (entry && now - entry.firstAt <= LOCK_WINDOW_MS && entry.count >= MAX_FAILS) {
    return res.status(429).json({
      ok: false,
      message: 'กรอกรหัส PIN ผิดเกินกำหนด กรุณารอ 60 วินาทีแล้วลองใหม่ค่ะ'
    });
  }

  const body = await readJsonBody(req);
  const pin = String((body && body.pin) || '').trim();
  if (pin.length < 4 || pin.length > 32) {
    return res.status(400).json({ ok: false, message: 'กรุณากรอกรหัส PIN ให้ถูกต้องค่ะ' });
  }

  const salt = process.env.ADMIN_PIN_SALT || DEFAULT_ADMIN_SALT;
  const expectedHex = process.env.ADMIN_PIN_HASH || DEFAULT_ADMIN_PIN_HASH;

  const actualBuf = crypto.createHash('sha256').update(salt + ':' + pin, 'utf8').digest();
  const expectedBuf = Buffer.from(expectedHex, 'hex');

  if (actualBuf.length === expectedBuf.length && crypto.timingSafeEqual(actualBuf, expectedBuf)) {
    attemptsByIp.delete(ip);
    return res.status(200).json({ ok: true, message: 'ยืนยันสิทธิ์แอดมินสำเร็จ' });
  }

  if (!entry || now - entry.firstAt > LOCK_WINDOW_MS) {
    attemptsByIp.set(ip, { count: 1, firstAt: now });
  } else {
    entry.count += 1;
  }

  return res.status(401).json({ ok: false, message: 'รหัส PIN แอดมินไม่ถูกต้อง' });
};
