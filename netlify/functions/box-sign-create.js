// netlify/functions/box-sign-complete.js
//
// Where Box sends the signer's browser after they finish signing.
// Marks the record as signed, then shows a confirmation page.
//
// Called as a GET redirect, e.g.
//   /.netlify/functions/box-sign-complete?report_id=<uuid>&doc_type=tm&signer=Doug
//
// Required env vars (already set for the archive job):
//   SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY

const SUPA_URL = process.env.SUPABASE_URL || 'https://uicmfyudiullulbbwzmh.supabase.co';
const SUPA_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const SITE_URL = process.env.URL || 'https://bespoke-sawine-3a49c3.netlify.app';

// doc_type → which table holds the record.
const TABLES = {
  daily: { table: 'daily_reports', label: 'Daily Report' },
  tm:    { table: 'tm_tickets',    label: 'T&M Ticket' },
};

function page({ ok, title, message, detail }) {
  const accent = ok ? '#34D399' : '#FBBF24';
  return `<!DOCTYPE html>
<html><head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>${title}</title>
<style>
  *{box-sizing:border-box}
  body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;
       background:#0D0D0F;color:#F0F4FF;font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;padding:20px}
  .card{background:#1A1A20;border:1px solid #26262E;border-left:4px solid ${accent};
        border-radius:16px;padding:36px 30px;max-width:440px;width:100%;text-align:center}
  .mark{font-size:52px;line-height:1;margin-bottom:14px}
  h1{font-size:21px;margin:0 0 10px;color:${accent}}
  p{font-size:14px;line-height:1.65;color:#8A94AD;margin:0 0 8px}
  .detail{font-size:12px;color:#5A6478;margin-top:14px}
  .brand{margin-top:26px;padding-top:18px;border-top:1px solid #26262E;font-size:11px;color:#5A6478}
  a{color:#60A5FA;text-decoration:none}
</style>
</head><body>
  <div class="card">
    <div class="mark">${ok ? '✅' : '⚠️'}</div>
    <h1>${title}</h1>
    <p>${message}</p>
    ${detail ? `<div class="detail">${detail}</div>` : ''}
    <div class="brand">
      AIME Field Pro<br/>
      Atlantic Industrial Mechanical &amp; Environmental Inc.
    </div>
  </div>
</body></html>`;
}

const html = (statusCode, body) => ({
  statusCode,
  headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' },
  body,
});

export const handler = async (event) => {
  const q = event.queryStringParameters || {};
  const reportId = q.report_id;
  const docType = q.doc_type || 'daily'; // legacy links carry no doc_type
  const signer = q.signer || 'the signer';

  const cfg = TABLES[docType] || TABLES.daily;

  // Always show the signer a friendly page — their signature is already
  // recorded in Box regardless of whether our own status update succeeds.
  if (!reportId) {
    return html(200, page({
      ok: true,
      title: 'Signature Received',
      message: 'Thank you — your signature has been recorded in Box.',
      detail: 'You may now close this window.',
    }));
  }

  try {
    if (!SUPA_KEY) throw new Error('SUPABASE_SERVICE_ROLE_KEY not set');

    const res = await fetch(
      `${SUPA_URL}/rest/v1/${cfg.table}?id=eq.${encodeURIComponent(reportId)}`,
      {
        method: 'PATCH',
        headers: {
          apikey: SUPA_KEY,
          Authorization: `Bearer ${SUPA_KEY}`,
          'Content-Type': 'application/json',
          Prefer: 'return=representation',
        },
        body: JSON.stringify({ hellosign_status: 'signed' }),
      }
    );

    if (!res.ok) throw new Error(`${res.status} ${await res.text()}`);

    const rows = await res.json();
    if (!Array.isArray(rows) || rows.length === 0) {
      // Signature is valid in Box; we just couldn't match a row.
      console.warn(`[box-sign-complete] no ${cfg.table} row for id ${reportId}`);
      return html(200, page({
        ok: true,
        title: 'Signature Received',
        message: `Thank you, ${signer}. Your signature has been recorded in Box.`,
        detail: 'You may now close this window.',
      }));
    }

    console.log(`[box-sign-complete] ${cfg.table} ${reportId} → signed`);
    return html(200, page({
      ok: true,
      title: 'Thank You — Signed',
      message: `Your signature on this ${cfg.label} has been received and recorded.`,
      detail: 'A copy has been saved to Box. You may now close this window.',
    }));
  } catch (e) {
    console.error('[box-sign-complete]', e.message);
    // Never show the signer a failure — the signature itself did succeed.
    return html(200, page({
      ok: true,
      title: 'Signature Received',
      message: `Thank you, ${signer}. Your signature has been recorded in Box.`,
      detail: 'You may now close this window.',
    }));
  }
};
