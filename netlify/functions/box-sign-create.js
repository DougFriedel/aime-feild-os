// netlify/functions/box-sign-create.js
// ESM build — use this version when the host repo's package.json has "type": "module".
// Creates a Box Sign signature request for a daily field report or a T&M ticket.
//
// Payload MUST include docType: "daily" | "tm".
// Requests without docType are treated as "daily" for backward compatibility.

import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';

const NETLIFY_URL = 'https://bespoke-sawine-3a49c3.netlify.app';

// Optional: set BOX_FOLDER_ID in Netlify env to keep signed docs out of the Box root.
const BOX_FOLDER_ID = process.env.BOX_FOLDER_ID || '0';

/* ── document type configuration ───────────────────────────── */
const DOC_TYPES = {
  daily: {
    label: 'Daily Report',
    title: 'Daily Field Report',
    folder: 'Daily Reports',
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
    title: 'Time & Materials Ticket',
    folder: 'T&M Tickets',
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

/* ── PDF rendering ─────────────────────────────────────────
   Builds a print-quality PDF of a T&M ticket / daily report using pdf-lib.
   Mirrors the layout of printTicket() in the app. */


const NAVY = rgb(0.122, 0.220, 0.392);   // #1F3864
const HEADROW = rgb(0.235, 0.353, 0.541); // column header blue
const GREY = rgb(0.42, 0.45, 0.50);
const LINE = rgb(0.78, 0.80, 0.84);
const SOFT = rgb(0.929, 0.941, 0.961);
const BLACK = rgb(0.1, 0.1, 0.12);
const WHITE = rgb(1, 1, 1);

const PW = 612, PH = 792, M = 36;
const CW = PW - M * 2;

const money = (n) => {
  const v = typeof n === 'string' ? parseFloat(String(n).replace(/[^0-9.-]/g, '')) : n;
  return '$' + Number(v || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};
const asMoney = (s) => (typeof s === 'string' && s.trim().startsWith('$') ? s : money(s));

async function buildTicketPdf(cfg, p) {
  const doc = await PDFDocument.create();
  const F = await doc.embedFont(StandardFonts.Helvetica);
  const FB = await doc.embedFont(StandardFonts.HelveticaBold);
  let page = doc.addPage([PW, PH]);
  let y = PH - M;

  const txt = (s, x, yy, size = 8, font = F, color = BLACK) =>
    page.drawText(String(s == null ? '' : s), { x, y: yy, size, font, color });

  const rightTxt = (s, xRight, yy, size = 8, font = F, color = BLACK) => {
    const w = font.widthOfTextAtSize(String(s ?? ''), size);
    txt(s, xRight - w, yy, size, font, color);
  };

  const clip = (s, font, size, maxW) => {
    s = String(s == null ? '' : s);
    if (font.widthOfTextAtSize(s, size) <= maxW) return s;
    while (s.length > 1 && font.widthOfTextAtSize(s + '…', size) > maxW) s = s.slice(0, -1);
    return s + '…';
  };

  const newPageIfNeeded = (need) => {
    if (y - need > M + 30) return;
    page = doc.addPage([PW, PH]);
    y = PH - M;
  };

  /* ── header ── */
  page.drawRectangle({ x: M, y: y - 26, width: 92, height: 26, color: NAVY });
  txt('AIME', M + 12, y - 19, 17, FB, WHITE);
  txt('ATLANTIC INDUSTRIAL MECHANICAL', M + 100, y - 9, 7.5, FB, NAVY);
  txt('& Environmental Inc.', M + 100, y - 17, 6.5, F, rgb(0.33, 0.33, 0.33));
  txt('5730 Pennington Ave, Baltimore, MD 21226', M + 100, y - 24, 6, F, GREY);

  rightTxt(cfg.title, PW - M, y - 12, 15, FB, BLACK);
  rightTxt(`${cfg.numberLabel} ${p.reportNo || '—'}`, PW - M, y - 24, 8, F, rgb(0.25, 0.25, 0.25));
  rightTxt(`Date: ${p.reportDate || '—'}`, PW - M, y - 34, 8, F, rgb(0.25, 0.25, 0.25));

  y -= 46;
  page.drawLine({ start: { x: M, y }, end: { x: PW - M, y }, thickness: 1.4, color: NAVY });
  y -= 12;

  /* ── info grid ── */
  const rows = [
    [['Project', p.projectName], ['Customer', p.customer]],
    [['PO #', p.poNumber], ['AFE/WO #', p.afeNumber]],
    [['Location', p.location], ['Submitted By', p.submittedBy]],
  ];
  const half = CW / 2;
  page.drawRectangle({ x: M, y: y - rows.length * 15, width: CW, height: rows.length * 15, borderColor: LINE, borderWidth: 0.7 });
  rows.forEach((pair, i) => {
    const ry = y - (i + 1) * 15 + 4;
    pair.forEach(([label, val], c) => {
      const x0 = M + c * half;
      txt(label, x0 + 5, ry, 6.5, FB, GREY);
      txt(clip(val || '—', F, 8, half - 78), x0 + 72, ry, 8, F, BLACK);
    });
    if (i < rows.length - 1)
      page.drawLine({ start: { x: M, y: y - (i + 1) * 15 }, end: { x: PW - M, y: y - (i + 1) * 15 }, thickness: 0.5, color: LINE });
  });
  page.drawLine({ start: { x: M + half, y }, end: { x: M + half, y: y - rows.length * 15 }, thickness: 0.5, color: LINE });
  y -= rows.length * 15 + 14;

  /* ── description ── */
  if (p.description) {
    newPageIfNeeded(40);
    txt('DESCRIPTION OF WORK', M, y, 6.5, FB, GREY);
    y -= 10;
    const words = String(p.description).split(/\s+/);
    let line = '';
    for (const w of words) {
      if (F.widthOfTextAtSize(line + ' ' + w, 8) > CW - 4) { txt(line.trim(), M + 2, y, 8); y -= 10; line = w; newPageIfNeeded(20); }
      else line += ' ' + w;
    }
    if (line.trim()) { txt(line.trim(), M + 2, y, 8); y -= 10; }
    y -= 6;
  }

  /* ── tables ── */
  const table = (title, cols, data, totalLabel, totalValue) => {
    if (!data || !data.length) return;
    newPageIfNeeded(28 + data.length * 13 + 16);

    // section bar
    page.drawRectangle({ x: M, y: y - 12, width: CW, height: 12, color: NAVY });
    txt(title.toUpperCase(), M + 5, y - 9, 7, FB, WHITE);
    y -= 12;

    // column header
    page.drawRectangle({ x: M, y: y - 12, width: CW, height: 12, color: HEADROW });
    let x = M;
    cols.forEach((c) => {
      if (c.align === 'right') rightTxt(c.label, x + c.w - 5, y - 8.5, 6.5, FB, WHITE);
      else if (c.align === 'center') {
        const tw = FB.widthOfTextAtSize(c.label, 6.5);
        txt(c.label, x + (c.w - tw) / 2, y - 8.5, 6.5, FB, WHITE);
      } else txt(c.label, x + 5, y - 8.5, 6.5, FB, WHITE);
      x += c.w;
    });
    y -= 12;

    // rows
    data.forEach((r) => {
      newPageIfNeeded(26);
      let cx = M;
      cols.forEach((c) => {
        const raw = r[c.key];
        const val = c.money ? asMoney(raw) : (raw == null || raw === '' ? '' : String(raw));
        if (c.align === 'right') rightTxt(clip(val, F, 7.5, c.w - 10), cx + c.w - 5, y - 9, 7.5);
        else if (c.align === 'center') {
          const tw = F.widthOfTextAtSize(clip(val, F, 7.5, c.w - 10), 7.5);
          txt(clip(val, F, 7.5, c.w - 10), cx + (c.w - tw) / 2, y - 9, 7.5);
        } else txt(clip(val, F, 7.5, c.w - 10), cx + 5, y - 9, 7.5);
        cx += c.w;
      });
      page.drawLine({ start: { x: M, y: y - 13 }, end: { x: PW - M, y: y - 13 }, thickness: 0.4, color: LINE });
      y -= 13;
    });

    // total strip
    if (totalLabel) {
      page.drawRectangle({ x: M, y: y - 13, width: CW, height: 13, color: SOFT });
      rightTxt(totalLabel, PW - M - 78, y - 9.5, 7.5, FB, BLACK);
      rightTxt(asMoney(totalValue), PW - M - 5, y - 9.5, 7.5, FB, BLACK);
      y -= 13;
    }
    y -= 10;
  };

  const it = p.lineItems || {};

  table('Labor',
    [{ key: 'name', label: 'Name', w: 150 }, { key: 'classification', label: 'Classification', w: 130 },
     { key: 'hours', label: 'Hours', w: 60, align: 'center' }, { key: 'rate', label: 'Rate/Hr', w: 110, align: 'right', money: true },
     { key: 'amount', label: 'Amount', w: 90, align: 'right', money: true }],
    it.labor, 'Labor Total', p.laborTotal);

  table('Equipment',
    [{ key: 'description', label: 'Equipment', w: 210 }, { key: 'unit', label: 'Unit', w: 80, align: 'center' },
     { key: 'qty', label: 'Qty', w: 60, align: 'center' }, { key: 'rate', label: 'Rate', w: 100, align: 'right', money: true },
     { key: 'amount', label: 'Amount', w: 90, align: 'right', money: true }],
    it.equipment, 'Equipment Total', p.equipmentTotal);

  table('Rental Equipment',
    [{ key: 'description', label: 'Description', w: 250 }, { key: 'qty', label: 'Qty', w: 90, align: 'center' },
     { key: 'rate', label: 'Rate', w: 100, align: 'right', money: true },
     { key: 'amount', label: 'Amount', w: 100, align: 'right', money: true }],
    it.rental, 'Rental Total', p.rentalTotal);

  table('Materials',
    [{ key: 'description', label: 'Description', w: 250 }, { key: 'qty', label: 'Qty', w: 70, align: 'center' },
     { key: 'unit', label: 'Unit', w: 60, align: 'center' },
     { key: 'unit_price', label: 'Unit Price', w: 90, align: 'right', money: true },
     { key: 'amount', label: 'Amount', w: 70, align: 'right', money: true }],
    it.materials, 'Materials Total', p.materialsTotal);

  table('Other Charges',
    [{ key: 'description', label: 'Description', w: 440 }, { key: 'amount', label: 'Amount', w: 100, align: 'right', money: true }],
    it.other, 'Other Total', p.otherTotal);

  /* ── summary box (right aligned, like the print layout) ── */
  const sum = [];
  if (p.laborTotal) sum.push(['Labor', p.laborTotal]);
  if (p.equipmentTotal) sum.push(['Equipment', p.equipmentTotal]);
  if (p.rentalTotal) sum.push(['Rental', p.rentalTotal]);
  if (p.materialsTotal) sum.push(['Materials', p.materialsTotal]);
  if (p.otherTotal) sum.push(['Other', p.otherTotal]);
  if (p.subtotal) sum.push(['Subtotal', p.subtotal]);
  if (p.markupAmount && parseFloat(p.markupPct) > 0) sum.push([`Markup (${p.markupPct}%)`, p.markupAmount]);

  newPageIfNeeded(sum.length * 13 + 60);
  const boxW = 220, boxX = PW - M - boxW;
  sum.forEach(([l, v], i) => {
    const ry = y - (i + 1) * 13;
    page.drawRectangle({ x: boxX, y: ry, width: boxW, height: 13, borderColor: LINE, borderWidth: 0.5 });
    txt(l, boxX + 6, ry + 3.5, 7.5);
    rightTxt(asMoney(v), PW - M - 6, ry + 3.5, 7.5);
  });
  y -= sum.length * 13;
  page.drawRectangle({ x: boxX, y: y - 18, width: boxW, height: 18, color: NAVY });
  txt('GRAND TOTAL', boxX + 6, y - 12.5, 9, FB, WHITE);
  rightTxt(asMoney(p.grandTotal), PW - M - 6, y - 12.5, 10, FB, WHITE);
  y -= 18 + 26;

  /* ── signature block ── */
  newPageIfNeeded(80);
  page.drawLine({ start: { x: M, y }, end: { x: PW - M, y }, thickness: 0.7, color: LINE });
  y -= 12;
  txt('PREPARED BY / PM', M, y, 6.5, FB, GREY);
  txt('CUSTOMER APPROVAL', M + CW / 2, y, 6.5, FB, GREY);
  y -= 8;

  // PM signature image if the app supplied one
  if (p.pmSignature && /^data:image\/(png|jpe?g);base64,/.test(p.pmSignature)) {
    try {
      const b64 = p.pmSignature.split(',')[1];
      const bytes = Uint8Array.from(Buffer.from(b64, 'base64'));
      const img = p.pmSignature.includes('image/png') ? await doc.embedPng(bytes) : await doc.embedJpg(bytes);
      const dims = img.scaleToFit(150, 34);
      page.drawImage(img, { x: M, y: y - dims.height, width: dims.width, height: dims.height });
    } catch (e) { /* fall through to a blank line */ }
  }
  page.drawLine({ start: { x: M, y: y - 38 }, end: { x: M + 170, y: y - 38 }, thickness: 0.7, color: BLACK });
  txt(p.pmName || p.submittedBy || '', M, y - 48, 8, FB);
  txt(`Date: ${p.reportDate || ''}`, M, y - 58, 7, F, GREY);

  const cx2 = M + CW / 2;
  page.drawLine({ start: { x: cx2, y: y - 38 }, end: { x: cx2 + 200, y: y - 38 }, thickness: 0.7, color: BLACK });
  txt('Name / Signature', cx2, y - 48, 7, F, GREY);
  page.drawLine({ start: { x: cx2, y: y - 66 }, end: { x: cx2 + 200, y: y - 66 }, thickness: 0.7, color: BLACK });
  txt('Date / Title', cx2, y - 76, 7, F, GREY);
  y -= 90;

  /* ── attestation + footer ── */
  newPageIfNeeded(40);
  cfg.attestation.forEach((l) => { txt(l, M, y, 7, F, rgb(0.3, 0.3, 0.34)); y -= 9; });

  const pages = doc.getPages();
  pages.forEach((pg, i) => {
    pg.drawText(
      `AIME Field Pro  ·  ${p.projectName || ''}  ·  ${cfg.numberLabel} ${p.reportNo || ''}  ·  Page ${i + 1} of ${pages.length}`,
      { x: M, y: 22, size: 6.5, font: F, color: GREY }
    );
  });

  return await doc.save();
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
    'Content-Type: application/pdf',
    '',
    '',
  ].join(CRLF);

  const body = Buffer.concat([
    Buffer.from(metaPart + CRLF + filePart, 'utf-8'),
    Buffer.from(content, 'utf-8'),
    Buffer.from(`${CRLF}--${boundary}--${CRLF}`, 'utf-8'),
  ]);

  let res = await fetch('https://upload.box.com/api/2.0/files/content', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': `multipart/form-data; boundary=${boundary}`,
    },
    body,
  });

  if (res.ok) return (await res.json()).entries[0].id;

  const data = await res.json().catch(() => ({}));

  // Re-sending the same ticket hits an existing filename. Upload a new version
  // of that file rather than failing — Box keeps the prior versions, and each
  // signed document Box produces is stored separately either way.
  const existingId = data?.context_info?.conflicts?.id;
  if (res.status === 409 && existingId) {
    const vBoundary = 'BoxV' + Date.now();
    const vMeta = [
      `--${vBoundary}`,
      'Content-Disposition: form-data; name="attributes"',
      'Content-Type: application/json',
      '',
      JSON.stringify({ name: filename }),
    ].join(CRLF);
    const vFile = [
      `--${vBoundary}`,
      `Content-Disposition: form-data; name="file"; filename="${filename}"`,
      'Content-Type: application/pdf',
      '',
      '',
    ].join(CRLF);
    const vBody = Buffer.concat([
      Buffer.from(vMeta + CRLF + vFile, 'utf-8'),
      Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf-8'),
      Buffer.from(`${CRLF}--${vBoundary}--${CRLF}`, 'utf-8'),
    ]);

    res = await fetch(`https://upload.box.com/api/2.0/files/${existingId}/content`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': `multipart/form-data; boundary=${vBoundary}`,
      },
      body: vBody,
    });
    if (res.ok) return (await res.json()).entries[0].id;
    throw new Error('Box new-version upload failed: ' + JSON.stringify(await res.json().catch(() => ({}))));
  }

  throw new Error('Box upload failed: ' + JSON.stringify(data));
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

    const pdfBytes = await buildTicketPdf(cfg, p);
    const docBuffer = Buffer.from(pdfBytes);
    const safeNo = String(reportNo || reportId).replace(/[^A-Za-z0-9_-]/g, '');
    const filename = `${cfg.filePrefix}-${safeNo}-${reportDate || 'draft'}.pdf`;

    const token = await getBoxToken();

    // Folder layout in Box:
    //   <BOX_FOLDER_ID or 'AIME Field Pro - Signatures'>/
    //       <Job Name>/
    //           AIME-TM-...txt
    //           AIME-DailyReport-...txt
    //
    // The job folder is created the first time anything is sent for that job
    // and reused for every document afterwards. ensureFolder is idempotent:
    // Box answers an existing folder with 409 + its id, which we reuse.
    let parentId = BOX_FOLDER_ID;
    if (!parentId || parentId === '0') {
      // Box Sign refuses root as a signature request's parent folder.
      parentId = await ensureFolder(token, 'AIME Field Pro - Signatures', '0');
    }

    const jobFolderName = (projectName || 'Unassigned Job')
      .replace(/[/\\<>:"|?*]/g, '-')
      .trim()
      .slice(0, 200) || 'Unassigned Job';

    // <parent>/<Job Name>/<Daily Reports | T&M Tickets>/
    // Both levels are created on first use and reused thereafter.
    const jobFolderId = await ensureFolder(token, jobFolderName, parentId);
    const signFolderId = await ensureFolder(token, cfg.folder, jobFolderId);

    const fileId = await uploadReportToBox(token, docBuffer, filename, signFolderId);

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
