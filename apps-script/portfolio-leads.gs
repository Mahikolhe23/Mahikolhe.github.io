const SHEET_ID = "1NfKmBqo9-ImHvmrBSm0Me0mb9AldOLje9p2WqOYhbEY";
const REQUIRED_HEADERS = [
  "Timestamp",
  "Name",
  "Email",
  "Inquiry Type",
  "Message",
  "Newsletter",
  "Active",
  "LastNewsletterSent"
];

function doGet(e) {
  return handleLead(e);
}

function doPost(e) {
  return handleLead(e);
}

function handleLead(e) {
  const params = e && e.parameter ? e.parameter : {};
  const payload = normalizePayload(params);

  if (payload.website) {
    return jsonResponse({ ok: true, skipped: true });
  }

  const validation = validatePayload(payload);
  if (!validation.ok) {
    return jsonResponse({ ok: false, error: validation.error });
  }

  const duplicateKey = makeDuplicateKey(payload);
  const cache = CacheService.getScriptCache();
  if (cache.get(duplicateKey)) {
    return jsonResponse({ ok: true, duplicate: true });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    const headers = ensureHeaders(sheet);
    const row = buildRow(headers, payload);

    sheet.appendRow(row);
    cache.put(duplicateKey, "1", 60);
    sendAcknowledgementEmail(payload);

    return jsonResponse({ ok: true });
  } catch (error) {
    return jsonResponse({ ok: false, error: "Unable to save enquiry." });
  } finally {
    lock.releaseLock();
  }
}

function normalizePayload(params) {
  const message = sanitizeText(params.message || "", 2000);
  const newsletter = params.newsletter === "Yes";

  return {
    name: sanitizeText(params.name || "", 120),
    email: sanitizeEmail(params.email || ""),
    message,
    website: sanitizeText(params.website || "", 120),
    newsletter,
    inquiryType: getInquiryType(Boolean(message), newsletter)
  };
}

function validatePayload(payload) {
  if (!payload.name) return { ok: false, error: "Name is required." };
  if (!isValidEmail(payload.email)) return { ok: false, error: "Valid email is required." };
  if (!payload.message && !payload.newsletter) {
    return { ok: false, error: "Message or newsletter subscription is required." };
  }
  return { ok: true };
}

function ensureHeaders(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), REQUIRED_HEADERS.length);
  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const hasAnyHeader = currentHeaders.some(Boolean);

  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, REQUIRED_HEADERS.length).setValues([REQUIRED_HEADERS]);
    return REQUIRED_HEADERS.slice();
  }

  const headers = currentHeaders.filter(Boolean);
  const missingHeaders = REQUIRED_HEADERS.filter((header) => !headers.includes(header));

  if (missingHeaders.length) {
    sheet.getRange(1, headers.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    return headers.concat(missingHeaders);
  }

  return headers;
}

function buildRow(headers, payload) {
  const values = {
    "Timestamp": new Date(),
    "Name": payload.name,
    "Email": payload.email,
    "Inquiry Type": payload.inquiryType,
    "Message": payload.message,
    "Newsletter": payload.newsletter ? "Yes" : "No",
    "Active": payload.newsletter ? "Yes" : "No",
    "LastNewsletterSent": ""
  };

  return headers.map((header) => Object.prototype.hasOwnProperty.call(values, header) ? values[header] : "");
}

function sendAcknowledgementEmail(payload) {
  const template = getEmailTemplate(payload);

  MailApp.sendEmail({
    to: payload.email,
    subject: template.subject,
    body: template.textBody,
    htmlBody: template.htmlBody,
    name: "Mahendra Kolhe",
    replyTo: "mkolhe23@gmail.com"
  });
}

function getEmailTemplate(payload) {
  const safeName = escapeHtml(payload.name);
  const safeMessage = escapeHtml(payload.message);

  if (payload.inquiryType === "Newsletter Only") {
    return {
      subject: "Welcome to Mahendra Kolhe's data engineering updates",
      textBody: [
        `Hi ${payload.name},`,
        "",
        "Thanks for subscribing. I will occasionally share useful notes on data engineering, Azure, Python, AI, automation, and projects I am building.",
        "",
        "If you ever want to unsubscribe, reply with Unsubscribe.",
        "",
        "Regards,",
        "Mahendra Kolhe"
      ].join("\n"),
      htmlBody: [
        `<p>Hi ${safeName},</p>`,
        "<p>Thanks for subscribing. I will occasionally share useful notes on data engineering, Azure, Python, AI, automation, and projects I am building.</p>",
        "<p>If you ever want to unsubscribe, reply with <strong>Unsubscribe</strong>.</p>",
        "<p>Regards,<br>Mahendra Kolhe</p>"
      ].join("")
    };
  }

  if (payload.inquiryType === "Enquiry + Newsletter") {
    return {
      subject: "Thanks for reaching out and subscribing",
      textBody: [
        `Hi ${payload.name},`,
        "",
        "Thanks for reaching out. I have received your note and will review it carefully.",
        "",
        "Your message:",
        payload.message,
        "",
        "You are also subscribed to occasional updates on data engineering, Azure, Python, AI, automation, and projects I am building.",
        "",
        "Regards,",
        "Mahendra Kolhe"
      ].join("\n"),
      htmlBody: [
        `<p>Hi ${safeName},</p>`,
        "<p>Thanks for reaching out. I have received your note and will review it carefully.</p>",
        `<p><strong>Your message:</strong><br>${safeMessage}</p>`,
        "<p>You are also subscribed to occasional updates on data engineering, Azure, Python, AI, automation, and projects I am building.</p>",
        "<p>Regards,<br>Mahendra Kolhe</p>"
      ].join("")
    };
  }

  return {
    subject: "Thanks for reaching out",
    textBody: [
      `Hi ${payload.name},`,
      "",
      "Thanks for your message. I have received it and will get back to you over email.",
      "",
      "Your message:",
      payload.message,
      "",
      "Regards,",
      "Mahendra Kolhe"
    ].join("\n"),
    htmlBody: [
      `<p>Hi ${safeName},</p>`,
      "<p>Thanks for your message. I have received it and will get back to you over email.</p>",
      `<p><strong>Your message:</strong><br>${safeMessage}</p>`,
      "<p>Regards,<br>Mahendra Kolhe</p>"
    ].join("")
  };
}

function getInquiryType(hasMessage, newsletter) {
  if (hasMessage && newsletter) return "Enquiry + Newsletter";
  if (newsletter) return "Newsletter Only";
  return "Enquiry Only";
}

function sanitizeText(value, maxLength) {
  return String(value)
    .replace(/[\u0000-\u001F\u007F]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLength);
}

function sanitizeEmail(value) {
  return String(value).trim().toLowerCase().slice(0, 254);
}

function isValidEmail(email) {
  const blockedDomains = ["example.com", "test.com", "mailinator.com", "tempmail.com", "10minutemail.com"];
  const domain = email.split("@")[1] || "";
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)
    && !email.includes("..")
    && !domain.startsWith("-")
    && !domain.endsWith("-")
    && !blockedDomains.includes(domain);
}

function makeDuplicateKey(payload) {
  const raw = [payload.email, payload.inquiryType, payload.message, payload.newsletter ? "Yes" : "No"].join("|");
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return Utilities.base64EncodeWebSafe(digest).slice(0, 80);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
