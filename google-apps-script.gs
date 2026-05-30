const CONFIG = {
  VERSION: "2026-05-30-filip-visual-order-v1",
  SPREADSHEET_ID: "1WELZBCrqCTuSgyA4OVVFgNxw-VlshYaFbIXmC4L7HSA",
  SHEET_NAME: "FilipObjednavky",
  SHOP_NAME: "FILIP - radosť z kvetín",
  SHOP_EMAIL: "info@flamengo-kvety.sk",
  SHOP_PHONE: "+421 903 444 655",
  SHOP_ADDRESS: "29. augusta 63, 974 01 Banská Bystrica",
  FALLBACK_WEB_URL: "https://www.flamengo-kvety.sk/",
  FALLBACK_HERO_IMAGE_URL: "https://images.unsplash.com/photo-1490750967868-88aa4486c946?auto=format&fit=crop&w=1200&q=80",
};

const HEADERS = [
  "ID",
  "Vytvorene",
  "Stav",
  "Styl kytice",
  "Rozpocet",
  "Prilezitost",
  "Datum dorucenia",
  "Cas dorucenia",
  "Meno prijemcu",
  "Telefon prijemcu",
  "Adresa dorucenia",
  "Text karticky",
  "Poznamka",
  "Meno zakaznika",
  "Email zakaznika",
  "Telefon zakaznika",
  "Zdroj",
  "Email stav",
  "Email chyba",
];

function doGet(e) {
  if (e && e.parameter && e.parameter.__order === "1") {
    return handleOrder(parseParameterPayload(e.parameter));
  }

  return jsonResponse({
    ok: true,
    service: CONFIG.SHOP_NAME + " objednavky",
    version: CONFIG.VERSION,
    spreadsheetIdSet: Boolean(CONFIG.SPREADSHEET_ID),
  });
}

function doPost(e) {
  try {
    const payload = parsePayload(e);
    return handleOrder(payload);
  } catch (error) {
    return errorResponse({
      ok: false,
      error: error.message,
    });
  }
}

function handleOrder(payload) {
  try {
    validatePayload(payload);

    const sheet = getOrdersSheet();
    const orderId = createOrderId();
    const createdAt = new Date();

    const row = [
      orderId,
      createdAt,
      "Nova",
      payload.bouquetStyle || "",
      payload.budget || "",
      payload.occasion || "",
      payload.deliveryDate || "",
      payload.deliveryWindow || "",
      payload.recipientName || "",
      payload.recipientPhone || "",
      payload.deliveryAddress || "",
      payload.cardMessage || "",
      payload.note || "",
      payload.customerName || "",
      payload.customerEmail || "",
      payload.customerPhone || "",
      payload.source || "",
      "Caka",
      "",
    ];

    sheet.appendRow(row);
    const rowIndex = sheet.getLastRow();
    const emailResult = sendOrderEmails(payload, orderId);
    sheet.getRange(rowIndex, 18, 1, 2).setValues([[
      emailResult.ok ? "Odoslane" : "Chyba",
      emailResult.error || "",
    ]]);

    return successResponse({
      ok: true,
      orderId,
      email: emailResult.ok ? "sent" : "failed",
      webUrl: getReturnUrl(payload.source),
      heroImageUrl: getHeroImageUrl(payload.source),
    });
  } catch (error) {
    return errorResponse({
      ok: false,
      error: error.message,
    });
  }
}

function setupSheet() {
  getOrdersSheet();
}

function testOrder() {
  const testPayload = {
    bouquetStyle: "Kytica pre pekný deň",
    budget: "30 - 45 €",
    occasion: "Test",
    deliveryDate: "2026-06-01",
    deliveryWindow: "Dohodnúť telefonicky",
    recipientName: "Test prijemca",
    recipientPhone: "0900000000",
    deliveryAddress: "29. augusta 63, Banska Bystrica",
    cardMessage: "Test karticky",
    note: "Toto je test z Apps Script editora.",
    customerName: "Test zakaznik",
    customerEmail: CONFIG.SHOP_EMAIL,
    customerPhone: "0900000000",
    consent: true,
    source: "Apps Script testOrder",
  };

  const result = doPost({
    postData: {
      contents: JSON.stringify(testPayload),
    },
  });

  Logger.log(result.getContent());
}

function testEmail() {
  MailApp.sendEmail({
    to: CONFIG.SHOP_EMAIL,
    subject: "Test emailu | " + CONFIG.SHOP_NAME,
    body: "Toto je test odosielania emailov z Apps Scriptu pre " + CONFIG.SHOP_NAME + ".",
    name: CONFIG.SHOP_NAME,
  });

  Logger.log("Test email odoslany na: " + CONFIG.SHOP_EMAIL);
}

function parsePayload(e) {
  if (e && e.parameter && Object.keys(e.parameter).length) {
    return parseParameterPayload(e.parameter);
  }

  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("Chybaju data objednavky.");
  }

  return JSON.parse(e.postData.contents);
}

function parseParameterPayload(parameters) {
  const payload = Object.assign({}, parameters);
  payload.consent = payload.consent === true || payload.consent === "true" || payload.consent === "on";
  return payload;
}

function validatePayload(payload) {
  const requiredFields = [
    "bouquetStyle",
    "budget",
    "deliveryDate",
    "deliveryWindow",
    "recipientName",
    "recipientPhone",
    "deliveryAddress",
    "customerName",
    "customerEmail",
    "customerPhone",
  ];

  requiredFields.forEach((field) => {
    if (!payload[field]) {
      throw new Error("Chyba povinne pole: " + field);
    }
  });

  if (!payload.consent) {
    throw new Error("Chyba suhlas so spracovanim udajov.");
  }
}

function getOrdersSheet() {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  }

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  const values = headerRange.getValues()[0];
  const missingHeaders = HEADERS.some((header, index) => values[index] !== header);

  if (missingHeaders) {
    headerRange.setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }

  return sheet;
}

function getSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error("Nie je nastavene SPREADSHEET_ID a skript nie je pripojeny ku Google Sheetu.");
  }

  return spreadsheet;
}

function createOrderId() {
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const suffix = Math.floor(Math.random() * 900 + 100);
  return "FL-" + date + "-" + suffix;
}

function sendOrderEmails(payload, orderId) {
  try {
    sendCustomerEmail(payload, orderId);
    sendShopEmail(payload, orderId);
    return {
      ok: true,
      error: "",
    };
  } catch (error) {
    Logger.log("Email error: " + error.message);
    return {
      ok: false,
      error: error.message,
    };
  }
}

function sendCustomerEmail(payload, orderId) {
  const subject = "Potvrdenie objednavky " + orderId + " | " + CONFIG.SHOP_NAME;
  const htmlBody = buildCustomerEmailHtml(payload, orderId);
  const body = [
    "Dobry den, " + payload.customerName + ",",
    "",
    "Dakujeme za objednavku kvetov. Objednavku sme prijali a budeme vas kontaktovat s potvrdenim detailov.",
    "",
    "Cislo objednavky: " + orderId,
    "Typ kytice: " + payload.bouquetStyle,
    "Rozpocet: " + payload.budget,
    "Datum dorucenia: " + payload.deliveryDate,
    "Cas dorucenia: " + payload.deliveryWindow,
    "Adresa dorucenia: " + payload.deliveryAddress,
    "",
    "Ak potrebujete nieco zmenit, kontaktujte nas:",
    CONFIG.SHOP_PHONE,
    CONFIG.SHOP_EMAIL,
    CONFIG.SHOP_ADDRESS,
    "",
    CONFIG.SHOP_NAME,
  ].join("\n");

  MailApp.sendEmail({
    to: payload.customerEmail,
    subject,
    body,
    htmlBody,
    name: CONFIG.SHOP_NAME,
    replyTo: CONFIG.SHOP_EMAIL,
  });
}

function sendShopEmail(payload, orderId) {
  const subject = "Nova objednavka kvetov " + orderId;
  const htmlBody = buildShopEmailHtml(payload, orderId);
  const body = [
    "Prisla nova objednavka z webu.",
    "",
    "Cislo objednavky: " + orderId,
    "Zakaznik: " + payload.customerName,
    "Email: " + payload.customerEmail,
    "Telefon: " + payload.customerPhone,
    "",
    "Prijemca: " + payload.recipientName,
    "Telefon prijemcu: " + payload.recipientPhone,
    "Adresa: " + payload.deliveryAddress,
    "",
    "Typ kytice: " + payload.bouquetStyle,
    "Rozpocet: " + payload.budget,
    "Prilezitost: " + (payload.occasion || "-"),
    "Datum dorucenia: " + payload.deliveryDate,
    "Cas dorucenia: " + payload.deliveryWindow,
    "",
    "Text karticky:",
    payload.cardMessage || "-",
    "",
    "Poznamka:",
    payload.note || "-",
  ].join("\n");

  MailApp.sendEmail({
    to: CONFIG.SHOP_EMAIL,
    subject,
    body,
    htmlBody,
    name: CONFIG.SHOP_NAME + " web",
    replyTo: payload.customerEmail,
  });
}

function buildCustomerEmailHtml(payload, orderId) {
  const heroImageUrl = getHeroImageUrl(payload.source);
  return buildEmailLayout({
    heroImageUrl,
    eyebrow: "Objednávka prijatá",
    title: "Ďakujeme, objednávku sme prijali.",
    intro:
      "Dobrý deň, " +
      escapeHtml(payload.customerName) +
      ", vaša objednávka je už v kvetinárstve. Čoskoro vás budeme kontaktovať s potvrdením dostupnosti, ceny a doručenia.",
    badge: "Číslo objednávky: " + escapeHtml(orderId),
    rows: [
      ["Typ objednávky", payload.bouquetStyle],
      ["Rozpočet", payload.budget],
      ["Dátum doručenia", payload.deliveryDate],
      ["Čas doručenia", payload.deliveryWindow],
      ["Príjemca", payload.recipientName],
      ["Adresa doručenia", payload.deliveryAddress],
      ["Text na kartičku", payload.cardMessage || "-"],
      ["Poznámka", payload.note || "-"],
    ],
    footer:
      "Ak potrebujete objednávku upraviť, odpovedzte na tento email alebo zavolajte na " +
      escapeHtml(CONFIG.SHOP_PHONE) +
      ".",
  });
}

function buildShopEmailHtml(payload, orderId) {
  const heroImageUrl = getHeroImageUrl(payload.source);
  return buildEmailLayout({
    heroImageUrl,
    eyebrow: "Nová objednávka z webu",
    title: "Prišla nová objednávka kvetov.",
    intro: "Objednávka je uložená v Google Sheete. Nižšie sú najdôležitejšie údaje pre vybavenie.",
    badge: "Číslo objednávky: " + escapeHtml(orderId),
    rows: [
      ["Zákazník", payload.customerName],
      ["Email zákazníka", payload.customerEmail],
      ["Telefón zákazníka", payload.customerPhone],
      ["Príjemca", payload.recipientName],
      ["Telefón príjemcu", payload.recipientPhone],
      ["Adresa doručenia", payload.deliveryAddress],
      ["Typ objednávky", payload.bouquetStyle],
      ["Rozpočet", payload.budget],
      ["Príležitosť", payload.occasion || "-"],
      ["Dátum doručenia", payload.deliveryDate],
      ["Čas doručenia", payload.deliveryWindow],
      ["Text na kartičku", payload.cardMessage || "-"],
      ["Poznámka", payload.note || "-"],
    ],
    footer: "Odpoveď na tento email pôjde priamo zákazníkovi.",
  });
}

function buildEmailLayout(data) {
  const rows = data.rows
    .map(function (row) {
      return (
        '<tr><td style="padding:12px 0;border-bottom:1px solid #dce8df;color:#5d6b65;font-size:13px;font-weight:700;text-transform:uppercase;">' +
        escapeHtml(row[0]) +
        '</td><td style="padding:12px 0;border-bottom:1px solid #dce8df;color:#071413;font-size:15px;font-weight:700;text-align:right;">' +
        escapeHtml(row[1]) +
        "</td></tr>"
      );
    })
    .join("");

  return (
    '<div style="margin:0;padding:0;background:#f5f8f2;color:#071413;font-family:Arial,sans-serif;">' +
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f5f8f2;padding:24px 0;"><tr><td align="center">' +
    '<table role="presentation" width="620" cellspacing="0" cellpadding="0" style="width:620px;max-width:calc(100% - 28px);background:#ffffff;border:1px solid #dce8df;border-radius:14px;overflow:hidden;box-shadow:0 18px 50px rgba(3,24,35,.12);">' +
    '<tr><td><img src="' +
    escapeHtml(data.heroImageUrl) +
    '" alt="FILIP - radosť z kvetín" width="620" style="display:block;width:100%;height:220px;object-fit:cover;border:0;"></td></tr>' +
    '<tr><td style="padding:30px 28px 26px;">' +
    '<div style="color:#04765f;font-size:12px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;margin-bottom:10px;">' +
    escapeHtml(data.eyebrow) +
    "</div>" +
    '<h1 style="margin:0 0 12px;font-size:30px;line-height:1.08;color:#071413;">' +
    escapeHtml(data.title) +
    "</h1>" +
    '<p style="margin:0 0 20px;color:#5d6b65;font-size:16px;line-height:1.55;">' +
    data.intro +
    "</p>" +
    '<div style="display:inline-block;margin:0 0 20px;padding:9px 12px;border-radius:999px;background:#e1f6ef;color:#04765f;font-weight:800;font-size:14px;">' +
    data.badge +
    "</div>" +
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">' +
    rows +
    "</table>" +
    '<p style="margin:22px 0 0;color:#5d6b65;font-size:14px;line-height:1.55;">' +
    data.footer +
    "</p>" +
    '<p style="margin:22px 0 0;color:#071413;font-size:14px;line-height:1.55;font-weight:800;">' +
    escapeHtml(CONFIG.SHOP_NAME) +
    '<br><span style="font-weight:500;color:#5d6b65;">' +
    escapeHtml(CONFIG.SHOP_ADDRESS) +
    "<br>" +
    escapeHtml(CONFIG.SHOP_PHONE) +
    " · " +
    escapeHtml(CONFIG.SHOP_EMAIL) +
    "</span></p>" +
    "</td></tr></table>" +
    "</td></tr></table></div>"
  );
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function successResponse(data) {
  const webUrl = getReturnUrl(data.webUrl);
  const heroImageUrl = data.heroImageUrl || CONFIG.FALLBACK_HERO_IMAGE_URL;
  const emailUrl = getEmailComposeUrl(CONFIG.SHOP_EMAIL, "Objednávka kvetov - " + CONFIG.SHOP_NAME);
  return HtmlService.createHtmlOutput(
    '<!doctype html><html lang="sk"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Objednávka prijatá</title><style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#f5f8f2;color:#071413;font-family:Inter,Arial,sans-serif}.box{width:min(640px,calc(100% - 36px));border:1px solid rgba(7,36,53,.16);border-radius:16px;background:white;box-shadow:0 24px 70px rgba(3,24,35,.12);overflow:hidden}.hero{height:230px;background:url(' +
      escapeCssUrl(heroImageUrl) +
      ') center/cover no-repeat}.content{padding:34px}.eyebrow{margin:0 0 10px;color:#04765f;font-size:12px;font-weight:900;letter-spacing:.08em;text-transform:uppercase}.id{display:inline-block;margin:0 0 18px;padding:8px 11px;border-radius:999px;background:#e1f6ef;color:#04765f;font-weight:800}h1{margin:0 0 12px;font-size:3.3rem;line-height:1}.lead{margin:0 0 22px;color:#5d6b65;line-height:1.6;font-size:16px}.actions{display:flex;flex-wrap:wrap;gap:10px}a{display:inline-flex;min-height:46px;align-items:center;justify-content:center;padding:0 16px;border-radius:8px;text-decoration:none;font-weight:800}.primary{background:#072435;color:white}.ghost{background:#e1f6ef;color:#04765f}@media(max-width:560px){.content{padding:26px}.hero{height:190px}h1{font-size:2.35rem}.actions a{width:100%}}</style></head><body><main class="box"><div class="hero" aria-hidden="true"></div><div class="content"><p class="eyebrow">FILIP - radosť z kvetín</p><span class="id">Číslo objednávky: ' +
      escapeHtml(data.orderId) +
      '</span><h1>Ďakujeme, objednávka je prijatá.</h1><p class="lead">Objednávku sme odoslali do kvetinárstva. Čoskoro vám ju potvrdíme emailom alebo telefonicky a doladíme prípadné detaily doručenia.</p><div class="actions"><a class="primary" href="' +
      escapeHtml(webUrl) +
      '">Späť na web</a><a class="ghost" target="_blank" rel="noreferrer" href="' +
      escapeHtml(emailUrl) +
      '">Napísať email</a></div></div></main></body></html>'
  );
}

function errorResponse(data) {
  return HtmlService.createHtmlOutput(
    '<!doctype html><html lang="sk"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Objednávku sa nepodarilo odoslať</title><style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#f5f8f2;color:#071413;font-family:Inter,Arial,sans-serif}.box{width:min(560px,calc(100% - 36px));padding:34px;border:1px solid rgba(163,59,48,.22);border-radius:12px;background:white;box-shadow:0 24px 70px rgba(3,24,35,.12)}h1{margin:0 0 12px;font-size:3rem;line-height:1}p{margin:0 0 20px;color:#5d6b65;line-height:1.55}.err{padding:12px;border-radius:8px;background:#fff1ef;color:#a33b30;font-weight:700}a{display:inline-flex;min-height:46px;align-items:center;justify-content:center;margin-top:18px;padding:0 16px;border-radius:8px;background:#072435;color:white;text-decoration:none;font-weight:800}@media(max-width:560px){h1{font-size:2.3rem}}</style></head><body><main class="box"><h1>Objednávku sa nepodarilo odoslať.</h1><p>Skúste to znova alebo zavolajte do kvetinárstva.</p><div class="err">' +
      escapeHtml(data.error || "Neznama chyba") +
      '</div><a href="' +
      escapeHtml(CONFIG.FALLBACK_WEB_URL) +
      '">Späť na web</a></main></body></html>'
  );
}

function getReturnUrl(source) {
  if (source && /^https?:\/\//i.test(source)) {
    return source;
  }

  return CONFIG.FALLBACK_WEB_URL;
}

function getHeroImageUrl(source) {
  const webUrl = getReturnUrl(source);

  if (/^https?:\/\//i.test(webUrl)) {
    const cleanUrl = webUrl.split("#")[0].split("?")[0];
    if (cleanUrl.endsWith("/")) {
      return cleanUrl + "assets/hero.webp";
    }
    return cleanUrl.replace(/\/[^\/]*$/, "/assets/hero.webp");
  }

  return CONFIG.FALLBACK_HERO_IMAGE_URL;
}

function getEmailComposeUrl(email, subject) {
  return "https://mail.google.com/mail/?view=cm&fs=1&to=" + encodeURIComponent(email) + "&su=" + encodeURIComponent(subject);
}

function escapeCssUrl(value) {
  return String(value).replace(/["'()\\\n\r]/g, "");
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
