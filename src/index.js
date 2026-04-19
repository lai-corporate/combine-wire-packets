import { PDFDocument, StandardFonts } from 'pdf-lib';
/ NOTE:
// This file is based on your current worker structure.
// It adds a new API route that accepts 2 PDFs + email HTML,
// merges them into one PDF, uploads that merged PDF to Flow,
// then updates NetSuite.
//
// IMPORTANT:
// This code assumes PDFDocument is available in the Worker runtime bundle.

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    try {
      if (request.method === 'GET' && url.pathname === '/favicon.ico') {
        return new Response(null, { status: 204 });
      }

      if (request.method === 'GET' && url.pathname === '/health') {
        return jsonResponse({ ok: true, service: 'wire-packet-worker' }, 200);
      }

      if (request.method === 'POST' && url.pathname === '/api/merge-wire-upload') {
        return handleMergeWireUpload(request, env);
      }

      return new Response('Not found', { status: 404 });
    } catch (err) {
      return jsonResponse({
        error: err && err.message ? err.message : String(err)
      }, 500);
    }
  }
};

async function handleMergeWireUpload(request, env) {
  try {
    validateApiKey(request, env);

    const body = await request.json().catch(() => null);
    if (!body) return jsonResponse({ error: 'Invalid JSON body.' }, 400);

    const bookingNumber = (body.bookingNumber || '').toString().trim();
    const billPaymentInternalId = (body.billPaymentInternalId || '').toString().trim();
    const billPaymentTranId = (body.billPaymentTranId || '').toString().trim();
    const outputFileName = (body.outputFileName || '').toString().trim();
    const email = body.email || {};
    const files = Array.isArray(body.files) ? body.files : [];

    if (!bookingNumber) return jsonResponse({ error: 'bookingNumber is required.' }, 400);
    if (!billPaymentTranId) return jsonResponse({ error: 'billPaymentTranId is required.' }, 400);
    if (!outputFileName) return jsonResponse({ error: 'outputFileName is required.' }, 400);

    const mainPacket = files.find(f => (f?.role || '') === 'mainPacket');
    const wirePacket = files.find(f => (f?.role || '') === 'wirePacket');

    if (!mainPacket?.fileContentBase64) {
      return jsonResponse({ error: 'mainPacket PDF is required.' }, 400);
    }
    if (!wirePacket?.fileContentBase64) {
      return jsonResponse({ error: 'wirePacket PDF is required.' }, 400);
    }

    const mergedPdfBytes = await mergeWirePacketToSinglePdf({
      bookingNumber,
      billPaymentInternalId,
      billPaymentTranId,
      email,
      mainPacketBase64: mainPacket.fileContentBase64,
      wirePacketBase64: wirePacket.fileContentBase64
    });

    const mergedBase64 = uint8ToBase64(mergedPdfBytes);

    const maxB64 = parseInt(env.MAX_MERGED_B64_CHARS || '35000000', 10);
    if (mergedBase64.length > maxB64) {
      return jsonResponse({
        error: 'Merged PDF too large for upload.',
        details: {
          base64Chars: mergedBase64.length,
          max: maxB64
        }
      }, 400);
    }

    const flow1 = await uploadMergedPdfToFlow(env, {
      bookingNumber,
      transactionNumber: billPaymentTranId,
      fileName: outputFileName,
      fileContentBase64: mergedBase64,
      mimeType: 'application/pdf',
      originalFileCount: 3,
      originalNames: [
        mainPacket.fileName || `${bookingNumber}_main.pdf`,
        wirePacket.fileName || `${bookingNumber}_wire.pdf`,
        `${bookingNumber}_${billPaymentTranId}_wire_email.pdf`
      ]
    });

    if (flow1.error) {
      return jsonResponse({
        error: 'Flow #1 (SharePoint upload) failed: ' + flow1.error,
        flow1
      }, 500);
    }

    const fileUrl = flow1.fileUrl || flow1.webUrl || flow1.sharepointUrl || null;
    if (!fileUrl) {
      return jsonResponse({
        error: 'Flow #1 succeeded but did not return a SharePoint file URL.',
        flow1
      }, 500);
    }

    const nsResult = await postToNetSuite(env, {
      billPaymentInternalId,
      bookingNumber,
      transactionNumber: billPaymentTranId,
      fileUrl
    });

    if (nsResult.error) {
      return jsonResponse({
        error: 'NetSuite POST failed: ' + nsResult.error,
        sharepoint: flow1,
        netsuite: nsResult
      }, 500);
    }

    return jsonResponse({
      ok: true,
      bookingNumber,
      billPaymentTranId,
      sharepoint: flow1,
      netsuite: nsResult
    });
  } catch (err) {
    return jsonResponse({ error: err.message || String(err) }, 500);
  }
}

function validateApiKey(request, env) {
  const configured = String(env.WORKER_API_KEY || '').trim();
  if (!configured) return;

  const provided = request.headers.get('x-worker-key') || '';
  if (provided !== configured) {
    throw new Error('Unauthorized');
  }
}

async function mergeWirePacketToSinglePdf(args) {
  const {
    bookingNumber,
    billPaymentInternalId,
    billPaymentTranId,
    email,
    mainPacketBase64,
    wirePacketBase64
  } = args;

  // PDFDocument must be available in runtime bundle
  const outPdf = await PDFDocument.create();

  // 1) main packet
  await appendPdfBase64(outPdf, mainPacketBase64);

  // 2) wire packet
  await appendPdfBase64(outPdf, wirePacketBase64);

  // 3) rendered email PDF
  const emailPdfBytes = await buildEmailPdf({
    bookingNumber,
    billPaymentInternalId,
    billPaymentTranId,
    email
  });
  const emailSrc = await PDFDocument.load(emailPdfBytes, { ignoreEncryption: true });
  const emailPages = await outPdf.copyPages(emailSrc, emailSrc.getPageIndices());
  emailPages.forEach(p => outPdf.addPage(p));

  return await outPdf.save();
}

async function appendPdfBase64(outPdf, base64) {
  const bytes = base64ToUint8Array(base64);
  const srcPdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
  const pageIndexes = srcPdf.getPageIndices();
  const pages = await outPdf.copyPages(srcPdf, pageIndexes);
  pages.forEach(p => outPdf.addPage(p));
}

async function buildEmailPdf({ bookingNumber, billPaymentInternalId, billPaymentTranId, email }) {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([612, 792]);

  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  let y = 750;
  const left = 42;
  const line = 16;

  function draw(text, opts = {}) {
    const size = opts.size || 10;
    const useFont = opts.bold ? bold : font;
    const lines = wrapText(String(text || ''), 85);

    for (const ln of lines) {
      if (y < 50) break;
      page.drawText(ln, { x: left, y, size, font: useFont });
      y -= line;
    }
    if (opts.gapAfter) y -= opts.gapAfter;
  }

  draw('WIRE EMAIL RECORD', { bold: true, size: 16, gapAfter: 8 });
  draw(`Booking Number: ${bookingNumber}`, { bold: true });
  draw(`Bill Payment Internal ID: ${billPaymentInternalId || ''}`);
  draw(`Transaction Number: ${billPaymentTranId}`, { bold: true, gapAfter: 6 });
  draw(`Subject: ${email.subject || ''}`, { bold: true });
  draw(`Received: ${email.receivedDateTime || ''}`);
  draw(`From: ${email.from || ''}`, { gapAfter: 8 });
  draw('EMAIL BODY', { bold: true, size: 12, gapAfter: 4 });
  draw(htmlToPlainText(email.bodyHtml || ''), { size: 9 });

  return await pdfDoc.save();
}

function wrapText(text, maxChars) {
  const words = String(text || '').replace(/\r/g, '').split(/\s+/);
  const lines = [];
  let current = '';

  for (const w of words) {
    const test = current ? current + ' ' + w : w;
    if (test.length <= maxChars) current = test;
    else {
      if (current) lines.push(current);
      current = w;
    }
  }
  if (current) lines.push(current);
  return lines.length ? lines : [''];
}

function htmlToPlainText(html) {
  return String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<\/(p|div|br|li|tr|h1|h2|h3|h4|h5|h6)>/gi, '\n')
    .replace(/<li>/gi, '• ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

// --- Flow #1: upload merged PDF ---
async function uploadMergedPdfToFlow(env, payload) {
  try {
    const flowUrl = env.PA_FLOW_UPLOAD_URL;
    if (!flowUrl) {
      return { error: 'PA_FLOW_UPLOAD_URL not configured in Worker settings.' };
    }

    const body = {
      bookingNumber: payload.bookingNumber,
      transactionNumber: payload.transactionNumber || '',
      file: {
        fileName: payload.fileName,
        mimeType: payload.mimeType || 'application/pdf',
        fileContentBase64: payload.fileContentBase64
      },
      originalFileCount: payload.originalFileCount,
      originalNames: payload.originalNames
    };

    const resp = await fetch(flowUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });

    const text = await resp.text();
    let json;
    try { json = JSON.parse(text); } catch { json = { raw: text }; }

    if (!resp.ok) {
      return { error: json.error || json.message || text, raw: json, status: resp.status };
    }
    if (json.success === false || json.error) {
      return { error: json.error || json.message || text, raw: json, status: resp.status };
    }

    return json;
  } catch (err) {
    return { error: err.message || String(err) };
  }
}

// --- Flow #2: NetSuite update ---
async function postToNetSuite(env, payload) {
  try {
    const flowUrl = env.PA_FLOW_NS_URL;
    if (!flowUrl) {
      return { error: 'PA_FLOW_NS_URL not configured in Worker settings.' };
    }

    const recordType = env.NS_RECORD_TYPE || 'vendorpayment';

    const body = {
      recordType,
      keyFieldId: 'internalid',
      keyValue: payload.billPaymentInternalId,
      fields: {
        custbody_wire_packet_done: true,
        custbody_wire_packet_file_url: payload.fileUrl,
        custbody_wire_packet_processed_dt: new Date().toISOString(),
        custbody_wire_packet_booking_no: payload.bookingNumber,
        custbody_wire_packet_status_note: `Merged main packet + wire packet + wire email for tranid ${payload.transactionNumber}`
      }
    };

    const resp = await fetch(flowUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });

    const text = await resp.text();
    let json;
    try { json = JSON.parse(text); } catch { json = { raw: text }; }

    if (!resp.ok) {
      return { error: json.error || json.message || text, raw: json, status: resp.status };
    }
    if (json.success === false || json.error) {
      return { error: json.error || json.message || text, raw: json, status: resp.status };
    }

    return json;
  } catch (err) {
    return { error: err.message || String(err) };
  }
}

function base64ToUint8Array(base64) {
  const binary = atob(String(base64 || '').trim());
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function uint8ToBase64(u8) {
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < u8.length; i += chunkSize) {
    const chunk = u8.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, chunk);
  }
  return btoa(binary);
}

function jsonResponse(obj, status = 200) {
  return new Response(JSON.stringify(obj, null, 2), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' }
  });
}
