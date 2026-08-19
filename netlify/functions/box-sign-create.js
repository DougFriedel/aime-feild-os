// netlify/functions/box-sign-create.js
// ESM build — use this version when the host repo's package.json has "type": "module".
// Creates a Box Sign signature request for a daily field report or a T&M ticket.
//
// Payload MUST include docType: "daily" | "tm".
// Requests without docType are treated as "daily" for backward compatibility.

const NETLIFY_URL = 'https://bespoke-sawine-3a49c3.netlify.app';

// Optional: set BOX_FOLDER_ID in Netlify env to keep signed docs out of the Box root.
const BOX_FOLDER_ID = process.env.BOX_FOLDER_ID || '0';

/* ── document type configuration ───────────────────────────── */
const DOC_TYPES = {
  daily: {
    label: 'Daily Report',
    heading: 'AIME Field Pro — Daily Report Sign-Off',
    numberLabel: 'Report #',
    filePrefix: 'AIME-DailyReport',
    subject: (no, proj) => `Sign Required — AIME Daily Report ${no || ''} · ${proj}`,
    message: (proj, date) =>
      `Please review and sign the daily field report for ${proj} dated ${date}.\n\nThank you,\nAIME Field Pro`,
    attestation: [
      'By signing below I confirm I have reviewed the work',
      'described in this daily field report and that it was',
      'performed satisfactorily.',
    ],
  },
  tm: {
    label: 'T&M Ticket',
    heading: 'AIME Field Pro — Time & Materials Ticket',
    numberLabel: 'T&M #',
    filePrefix: 'AIME-TM',
    subject: (no, proj) => `Signature Required — AIME T&M Ticket ${no || ''} · ${proj}`,
    message: (proj, date) =>
      `Please review and sign the Time & Materials ticket for ${proj} dated ${date}.\n\n` +
      `This ticket itemizes labor, equipment, and materials furnished on the date shown. ` +
      `Your signature authorizes these charges for invoicing.\n\nThank you,\nAIME Field Pro`,
    attestation: [
      'By signing below I confirm that the labor, equipment, and',
      'materials itemized above were furnished as described, that',
      'the hours and quantities shown are accurate, and that these',
      'charges are authorized for invoicing.',
    ],
  },
};

/* ── plain-text table rendering ────────────────────────────── */
const money = (n) =>
  '$' + Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// Renders a fixed-width table. cols: [{ key, label, width, align }]
function renderTable(title, cols, rows, totalLabel, totalValue) {
  if (!rows || !rows.length) return [];

  const pad = (s, w, align) => {
    s = String(s == null ? '' : s);
    if (s.length > w) s = s.slice(0, w - 1) + '…';
    return align === 'right' ? s.padStart(w) : align === 'center'
      ? s.padStart(Math.floor((w + s.length) / 2)).padEnd(w)
      : s.padEnd(w);
  };

  const line = cols.map((c) => '-'.repeat(c.width)).join('-+-');
  const header = cols.map((c) => pad(c.label, c.width, c.align)).join(' | ');

  const out = ['', title.toUpperCase(), header, line];
  rows.forEach((r) => {
    out.push(cols.map((c) => pad(r[c.key], c.width, c.align)).join(' | '));
  });

  if (totalLabel) {
    out.push(line);
    const totalWidth = cols.slice(0, -1).reduce((s, c) => s + c.width, 0) + (cols.length - 2) * 3;
    const lastCol = cols[cols.length - 1];
    out.push(pad(totalLabel, totalWidth, 'right') + ' | ' + pad(totalValue, lastCol.width, 'right'));
  }
  return out;
}

function buildDocument(cfg, p) {
  const rule = '='.repeat(72);
  const lines = [
    cfg.heading,
    rule,
    '',
    `Project:        ${p.projectName || '—'}`,
    `Customer:       ${p.customer || '—'}`,
    `${(cfg.numberLabel + ':').padEnd(16)}${p.reportNo || '—'}`,
    `Date:           ${p.reportDate || '—'}`,
    `Submitted By:   ${p.submittedBy || '—'}`,
  ];

  if (p.poNumber) lines.push(`PO #:           ${p.poNumber}`);
  if (p.afeNumber) lines.push(`AFE / WO #:     ${p.afeNumber}`);
  if (p.location) lines.push(`Location:       ${p.location}`);

  if (p.description) {
    lines.push('', 'DESCRIPTION OF WORK', '-'.repeat(72));
    // wrap at 72 chars
    String(p.description).split('\n').forEach((para) => {
      let cur = '';
      para.split(/\s+/).forEach((w) => {
        if ((cur + ' ' + w).trim().length > 72) { lines.push(cur.trim()); cur = w; }
        else cur += ' ' + w;
      });
      if (cur.trim()) lines.push(cur.trim());
    });
  }

  const items = p.lineItems || {};

  lines.push(
    ...renderTable(
      'Labor',
      [
        { key: 'name', label: 'Name', width: 20, align: 'left' },
        { key: 'classification', label: 'Classification', width: 18, align: 'left' },
        { key: 'hours', label: 'Hours', width: 7, align: 'center' },
        { key: 'rate', label: 'Rate', width: 10, align: 'right' },
        { key: 'amount', label: 'Amount', width: 11, align: 'right' },
      ],
      items.labor,
      'Labor Total',
      p.laborTotal
    )
  );

  lines.push(
    ...renderTable(
      'Equipment',
      [
        { key: 'description', label: 'Equipment', width: 30, align: 'left' },
        { key: 'unit', label: 'Unit', width: 8, align: 'center' },
        { key: 'qty', label: 'Qty', width: 7, align: 'center' },
        { key: 'rate', label: 'Rate', width: 10, align: 'right' },
        { key: 'amount', label: 'Amount', width: 11, align: 'right' },
      ],
      items.equipment,
      'Equipment Total',
      p.equipmentTotal
    )
  );

  lines.push(
    ...renderTable(
      'Rental Equipment',
      [
        { key: 'description', label: 'Description', width: 34, align: 'left' },
        { key: 'qty', label: 'Qty', width: 8, align: 'center' },
        { key: 'rate', label: 'Rate', width: 12, align: 'right' },
        { key: 'amount', label: 'Amount', width: 12, align: 'right' },
      ],
      items.rental,
      'Rental Total',
      p.rentalTotal
    )
  );

  lines.push(
    ...renderTable(
      'Materials',
      [
        { key: 'description', label: 'Description', width: 34, align: 'left' },
        { key: 'qty', label: 'Qty', width: 8, align: 'center' },
        { key: 'unit_price', label: 'Unit Price', width: 12, align: 'right' },
        { key: 'amount', label: 'Amount', width: 12, align: 'right' },
      ],
      items.materials,
      'Materials Total',
      p.materialsTotal
    )
  );

  lines.push(
    ...renderTable(
      'Other Charges',
      [
        { key: 'description', label: 'Description', width: 54, align: 'left' },
        { key: 'amount', label: 'Amount', width: 12, align: 'right' },
      ],
      items.other,
      'Other Total',
      p.otherTotal
    )
  );

  // Summary
  lines.push('', 'SUMMARY', '-'.repeat(72));
  const sum = (l, v) => lines.push(l.padEnd(56) + String(v).padStart(16));
  if (p.laborTotal) sum('Labor', p.laborTotal);
  if (p.equipmentTotal) sum('Equipment', p.equipmentTotal);
  if (p.rentalTotal) sum('Rental Equipment', p.rentalTotal);
  if (p.materialsTotal) sum('Materials', p.materialsTotal);
  if (p.otherTotal) sum('Other', p.otherTotal);
  if (p.subtotal) sum('Subtotal', p.subtotal);
  if (p.markupAmount && parseFloat(p.markupPct) > 0) sum(`Markup (${p.markupPct}%)`, p.markupAmount);
  lines.push('-'.repeat(72));
  sum('GRAND TOTAL', p.grandTotal || money(0));

  lines.push('', rule, '', ...cfg.attestation, '');
  return lines.join('\n');
}

/* ── Box helpers ───────────────────────────────────────────── */
async function getBoxToken() {
  const res = await fetch('https://api.box.com/oauth2/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: process.env.BOX_CLIENT_ID,
      client_secret: process.env.BOX_CLIENT_SECRET,
      grant_type: 'client_credentials',
      box_subject_type: 'enterprise',
      box_subject_id: process.env.BOX_ENTERPRISE_ID,
    }),
  });
  const data = await res.json();
  if (!data.access_token) throw new Error('Box auth failed: ' + JSON.stringify(data));
  return data.access_token;
}

// Box Sign refuses to use the root folder as a signature request's parent, so
// make sure a real subfolder exists and use that. Idempotent: if it's already
// there Box returns 409 with the existing id.
async function ensureFolder(token, name, parentId) {
  const res = await fetch('https://api.box.com/2.0/folders', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, parent: { id: String(parentId) } }),
  });
  if (res.ok) return (await res.json()).id;
  const err = await res.json().catch(() => ({}));
  const existing = err?.context_info?.conflicts?.[0]?.id;
  if (res.status === 409 && existing) return existing;
  throw new Error(`Box folder "${name}": ${res.status} ${JSON.stringify(err)}`);
}

async function uploadReportToBox(token, content, filename, parentId) {
  const boundary = 'BoxBoundary' + Date.now();
  const CRLF = '\r\n';

  const metaPart = [
    `--${boundary}`,
    'Content-Disposition: form-data; name="attributes"',
    'Content-Type: application/json',
    '',
    JSON.stringify({ name: filename, parent: { id: String(parentId) } }),
  ].join(CRLF);

  const filePart = [
    `--${boundary}`,
    `Content-Disposition: form-data; name="file"; filename="${filename}"`,
    'Content-Type: text/plain',
    '',
    '',
  ].join(CRLF);

  const body = Buffer.concat([
    Buffer.from(metaPart + CRLF + filePart, 'utf-8'),
    Buffer.from(content, 'utf-8'),
    Buffer.from(`${CRLF}--${boundary}--${CRLF}`, 'utf-8'),
  ]);

  const res = await fetch('https://upload.box.com/api/2.0/files/content', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': `multipart/form-data; boundary=${boundary}`,
    },
    body,
  });

  const data = await res.json();
  if (!res.ok) throw new Error('Box upload failed: ' + JSON.stringify(data));
  return data.entries[0].id;
}

/* ── handler ───────────────────────────────────────────────── */
export const handler = async (event) => {
  const corsHeaders = {
    'Access-Control-Allow-Origin': NETLIFY_URL,
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 200, headers: corsHeaders, body: '' };
  if (event.httpMethod !== 'POST')
    return { statusCode: 405, headers: corsHeaders, body: JSON.stringify({ error: 'Method not allowed' }) };

  try {
    const p = JSON.parse(event.body);
    const {
      docType = 'daily',
      reportId,
      inspectorEmail,
      inspectorName,
      projectName,
      reportNo,
      reportDate,
    } = p;

    const cfg = DOC_TYPES[docType];
    if (!cfg) throw new Error(`Unknown docType: ${docType}`);

    if (!process.env.BOX_CLIENT_ID) throw new Error('BOX_CLIENT_ID not set');
    if (!inspectorEmail) throw new Error('Signer email is required');
    if (!reportId) throw new Error('reportId is required');

    const docText = buildDocument(cfg, p);
    const safeNo = String(reportNo || reportId).replace(/[^A-Za-z0-9_-]/g, '');
    const filename = `${cfg.filePrefix}-${safeNo}-${reportDate || 'draft'}.txt`;

    const token = await getBoxToken();

    // If BOX_FOLDER_ID is unset (or root), create/reuse a real subfolder —
    // Box Sign rejects root as a parent folder.
    let signFolderId = BOX_FOLDER_ID;
    if (!signFolderId || signFolderId === '0') {
      signFolderId = await ensureFolder(token, 'AIME Field Pro - Signatures', '0');
    }

    const fileId = await uploadReportToBox(token, docText, filename, signFolderId);

    // doc_type is passed through so box-sign-complete knows which table to update
    const redirectUrl =
      `${NETLIFY_URL}/.netlify/functions/box-sign-complete` +
      `?report_id=${encodeURIComponent(reportId)}` +
      `&doc_type=${encodeURIComponent(docType)}` +
      `&signer=${encodeURIComponent(inspectorName || 'Signer')}`;

    const signRes = await fetch('https://api.box.com/2.0/sign_requests', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        signers: [{ email: inspectorEmail, name: inspectorName || 'Signer', role: 'signer' }],
        source_files: [{ id: fileId, type: 'file' }],
        parent_folder: { id: String(signFolderId), type: 'folder' },
        redirect_url: redirectUrl,
        declined_redirect_url: NETLIFY_URL,
        email_subject: cfg.subject(reportNo, projectName),
        email_message: cfg.message(projectName, reportDate),
        are_reminders_enabled: true,
        days_valid: 30,
        external_id: `${docType}:${reportId}`,
      }),
    });

    const signData = await signRes.json();
    if (!signRes.ok) throw new Error('Box Sign API error: ' + JSON.stringify(signData));

    return {
      statusCode: 200,
      headers: corsHeaders,
      body: JSON.stringify({ requestId: signData.id, status: 'pending', fileId, docType }),
    };
  } catch (err) {
    console.error('box-sign-create error:', err.message);
    return {
      statusCode: 500,
      headers: corsHeaders,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
