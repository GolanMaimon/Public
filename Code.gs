const SHEET_NAME = "טופסים";
const FOLDER_ID = "1o8d6ocYdAf29v7W6KlEk6pzvZfIlaRx8";
const LOG_SHEET = "Logs";

function logMsg(runId, msg) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let s = ss.getSheetByName(LOG_SHEET);
    if (!s) s = ss.insertSheet(LOG_SHEET);
    s.appendRow([new Date().toLocaleString("he-IL"), runId, msg]);
  } catch (e) {
    Logger.log("LOG ERROR: " + e.message);
  }
}

function doPost(e) {
  const runId = "RUN-" + new Date().getTime();
  try {
    logMsg(runId, "doPost start");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    logMsg(runId, "sheet found: " + Boolean(sheet));

    // יצירת גיליון וכותרות (תואמות בדיוק לתבנית שלך)
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "תאריך הגשה",
        "תאריך האירוע",
        "שם התלמיד/ה",
        "כיתה",
        "ת.ז. תלמיד/ה",
        "שם ההורה המלווה",
        "ת.ז. הורה מלווה",
        "כתובת",
        "טלפון נייד",
        "טלפון נוסף",
        "תאריך חתימה",
        "קישור לחתימה",
        "תמונה",
        "קישור ל-PDF"
      ]);
      sheet.getRange(1, 1, 1, 14)
        .setBackground("#0f2744")
        .setFontColor("white")
        .setFontWeight("bold");
      sheet.setFrozenRows(1);
      logMsg(runId, "sheet created");
    }

    const data = JSON.parse(e.postData.contents || "{}");
    logMsg(runId, "payload keys: " + Object.keys(data || {}).join(","));

    const childId = (data.childId || "").trim();
    const eventDate = (data.eventDate || "").trim();
    logMsg(runId, "childId=" + childId + ", eventDate=" + eventDate);

    // בדיקת כפילות
    const lastRow = sheet.getLastRow();
    if (childId && eventDate && lastRow > 1) {
      // עמודת ת.ז תלמיד היא עמודה מס' 5 (E)
      const ids = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
      const dates = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === childId && String(dates[i][0]).trim() === eventDate) {
          logMsg(runId, "duplicate found at row " + (i + 2));
          return ContentService
            .createTextOutput(JSON.stringify({ success: false, error: "כבר קיימת הרשמה לת.ז. זו בתאריך האירוע" }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    // שמירת חתימה כ-PNG
    let sigLink = "";
    let sigBlob = null;
    if (data.signature && String(data.signature).startsWith("data:image")) {
      logMsg(runId, "signature detected");
      const base64 = data.signature.split(",")[1];
      const bytes = Utilities.base64Decode(base64);
      sigBlob = Utilities.newBlob(bytes, "image/png", "signature.png");

      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file = folder.createFile(sigBlob.copyBlob())
        .setName(`signature_${childId || "unknown"}_${Date.now()}.png`);
      sigLink = file.getUrl();
      logMsg(runId, "signature saved: " + sigLink);
    } else {
      logMsg(runId, "no signature in payload");
    }

    // יצירת PDF
    const now = new Date();
    const pdfName = buildPdfName(data.childName || "תלמיד", now);
    logMsg(runId, "pdfName=" + pdfName);

    const pdfBlob = buildPdfBlob(data, sigBlob);
    logMsg(runId, "pdfBlob created, size=" + pdfBlob.getBytes().length);

    const folder = DriveApp.getFolderById(FOLDER_ID);
    const pdfFile = folder.createFile(pdfBlob).setName(pdfName);
    const pdfLink = pdfFile.getUrl();
    logMsg(runId, "pdf saved: " + pdfLink);

    // הוספת שורה לגיליון - הסדר תוקן שיתאים בדיוק לעמודות
    sheet.appendRow([
      now.toLocaleString("he-IL"),  // תאריך הגשה
      eventDate,                    // תאריך האירוע
      data.childName  || "",        // שם התלמיד/ה
      data.childClass || "",        // כיתה  <-- נוספה כאן בדיוק לאחר שם התלמיד לפי העמודות
      childId,                      // ת.ז. תלמיד/ה
      data.parentName || "",        // שם ההורה המלווה
      data.parentId   || "",        // ת.ז. הורה מלווה
      data.address    || "",        // כתובת
      data.phone      || "",        // טלפון נייד
      data.phone2     || "",        // טלפון נוסף
      data.signDate   || "",        // תאריך חתימה
      sigLink,                      // קישור לחתימה
      "",                           // תמונה
      pdfLink                       // קישור ל-PDF
    ]);

    const newLast = sheet.getLastRow();
    
    // חישוב הנוסחה לתמונה. מכיוון שהוספנו את 'כיתה', קישור חתימה נמצא בעמודה 12 (L), ועמודת התמונה היא 13 (M)
    sheet.getRange(newLast, 13)
      .setFormula('=IF(L' + newLast + '<>"",IMAGE(L' + newLast + '),"")');
    sheet.setRowHeight(newLast, 120);

    if (newLast % 2 === 0) {
      sheet.getRange(newLast, 1, 1, 14).setBackground("#e8f4fd");
    }

    logMsg(runId, "doPost success row=" + newLast);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, row: newLast }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    logMsg(runId, "ERROR: " + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function buildPdfName(childName, dateObj) {
  const safe = String(childName || "תלמיד")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^\u0590-\u05FFa-zA-Z0-9_]/g, "");
  const dd = String(dateObj.getDate()).padStart(2, "0");
  const mm = String(dateObj.getMonth() + 1).padStart(2, "0");
  const yy = String(dateObj.getFullYear()).slice(-2);
  const hh = String(dateObj.getHours()).padStart(2, "0");
  const mi = String(dateObj.getMinutes()).padStart(2, "0");
  return `${safe}_approve_${dd}-${mm}-${yy}_${hh}-${mi}.pdf`;
}

function applyRtl(paragraph) {
  const rtl = (DocumentApp.TextDirection && DocumentApp.TextDirection.RIGHT_TO_LEFT) || null;
  if (rtl) {
    paragraph.setTextDirection(rtl);
  }
  if (paragraph.setAlignment) {
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  }
  return paragraph;
}

function addRtlNumbered(body, index, text) {
  const rlm = "\u200f";
  const p = body.appendParagraph(`${rlm}${index}. ${text}`);
  applyRtl(p);
  return p;
}



function buildPdfBlob(data, sigBlob) {
  const yesNo = (v) => v === "yes" ? "כן" : v === "no" ? "לא" : "";
  const decls = Array.isArray(data.declarations) ? data.declarations : [];
  const declMap = {};
  decls.forEach(d => { declMap[d.id] = yesNo(d.value); });

  const doc = DocumentApp.create("temp_pdf");
  const body = doc.getBody();
  
  // הצרת שוליים כדי להכניס הכל לעמוד אחד מאורגן
  body.setMarginTop(20).setMarginBottom(20).setMarginLeft(20).setMarginRight(20);
  
  // הגדרת פונט קטן יותר וצפיפות כדי לוודא שזה נכנס לדף אחד
  const style = {};
  style[DocumentApp.Attribute.FONT_SIZE] = 9;
  style[DocumentApp.Attribute.LINE_SPACING] = 1.0;
  body.setAttributes(style);

  const rtl = (DocumentApp.TextDirection && DocumentApp.TextDirection.RIGHT_TO_LEFT) || null;
  if (rtl) body.setTextDirection(rtl);

  const title = body.appendParagraph("טופס אישור השתתפות במסיבת סיום שכבת כיתות ב׳")
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  if (rtl) title.setTextDirection(rtl);

  const intro = body.appendParagraph(
    'הורים ותלמידים יקרים, לקראת מסיבת הסיום השנתית של שכבת כיתות ב׳ (כיתות ב׳1 ו-ב׳2) ' +
    'מבית ספר "ערמונים", אשר תתקיים בתאריך 10/6/2026, בין השעות 16:00-19:00, במתחם בריכת כפר מעש, ' +
    'אנו מבקשים מכם לקרוא בעיון רב מסמך זה, לחתום עליו ולהחזירו לנציגי הוועד.'
  );
  applyRtl(intro);

  const notice = body.appendParagraph(
    "לתשומת לבכם: השתתפותו/ה של התלמיד/ה באירוע מותנית בהמצאת טופס זה כשהוא מלא וחתום במלואו, וכן בליווי הורה כפי שיפורט להלן."
  ).setBold(true);
  applyRtl(notice);

  const items = [
    "לתשומת לבכם, השתתפותו/ה של התלמיד/ה באירוע מותנית בהמצאת טופס זה כשהוא מלא וחתום במלואו, וכן בליווי הורה כפי שיפורט להלן.",
    "\"הוועד\" או \"נציגי הוועד\" הם הורים בבית ספר \"ערמונים\", הפועלים בהתנדבות מלאה.",
    "מובהר בזאת, כי תפקידם של נציגי הוועד מוגבל אך ורק לתיאום הלוגיסטי והטכני המקדים של מסיבת הסיום. הוועד ו/או מי מהורי הוועד אינם משמשים כמפיקים, מפעילים, משגיחים, מצילים או אחראים על ביטחונם של הילדים המשתתפים במהלך האירוע עצמו.",
    "חלה חובה מוחלטת על נוכחות פיזית של הורה (או אפוטרופוס חוקי) מלווה עבור כל ילד וילדה לאורך כל שעות הפעילות במסיבת הסיום במתחם הבריכה.",
    "לא תותר כניסה או שהות של ילד או ילדה למתחם הבריכה ללא השגחה צמודה ורציפה של ההורה שלו/ה.",
    "ההורה המלווה נושא באחריות המלאה להשגחה על ילדו/ה, בדגש על שהות במים ובסביבת הבריכה.",
    "ההורים מצהירים ומאשרים כי בדקו באופן אישי את מתחם האירוע, סביבת הבריכה והמתקנים בו, ומצאו כי השטח תקין, בטוח ומתאים לפעילות המתוכננת. בחירה: " + (declMap.d4 || ""),
    "מובהר ומוסכם בזאת כי אין ולא תהיה כל אחריות במישרין ו/או בעקיפין להורי ועד כיתות ב׳ בבי״ס ערמונים לכל נזק (גוף, נפש או רכוש), פציעה, אובדן או פגיעה שיגרמו לילד/ה או לרכושו/ה במהלך האירוע, בדרך אליו או בחזרה ממנו.",
    "מוסכם בזאת, כי האחריות המלאה, הבלעדית והמוחלטת לשלומו, ביטחונו, בריאותו והתנהגותו של הילד/ה מוטלת על ההורים בלבד, ולא על נציגי הוועד.",
    "מובהר ומוסכם, כי בכפוף להשגחת ההורה, כל ילד/ה נושא/ת באחריות אישית על התנהגותו/ה במהלך האירוע, על כל המשתמע מכך, ובמנותק לחלוטין מפעילות הוועד.",
    "ההורים מצהירים בזאת כי בנם/בתם יודע/ת לשחות בצורה טובה ומשביעת רצון, ומורשה/ית להיכנס למים... בחירה: " + (declMap.d1 || ""),
    "ההורים מתחייבים כי שוחחו עם הילד/ה והנחו אותו/ה להישמע... בחירה: " + (declMap.d3 || ""),
    "ההורים מצהירים כי הילד/ה כשיר/ה מבחינה בריאותית... בחירה: " + (declMap.d2 || ""),
    "מובהר ומוסכם בזאת, כי ההורים מוותרים ויתור סופי... בחירה: " + (declMap.d5 || ""),
    "למען הסר ספק מובהר בזאת, ככל שתוגש תביעה או דרישה בניגוד לאמור במסמך זה, מתחייבים ההורים החתומים מטה לשאת בכל ההוצאות והנזקים שיגרמו לוועד או לחבריו בעקבות זאת."
  ];

  items.forEach((t, i) => {
    addRtlNumbered(body, i + 1, t);
  });

  const confirm = body.appendParagraph(
    "אני החתום/ה מטה, הורה ו/או אפוטרופוס חוקי של התלמיד/ה, מאשר/ת בזאת כי קראתי בעיון את כל הסעיפים..."
  ).setBold(true);
  applyRtl(confirm);

  // קיבוץ כל פרטי הטופס לפסקה אחת מרוכזת כדי לחסוך מקום
  const infoText = 
    `• תלמיד/ה: ${data.childName || ""} | כיתה: ${data.childClass || ""} | ת.ז תלמיד: ${data.childId || ""}\n` +
    `• הורה מלווה: ${data.parentName || ""} | ת.ז הורה: ${data.parentId || ""}\n` +
    `• טלפון: ${data.phone || ""} | חירום: ${data.phone2 || ""} | כתובת: ${data.address || ""}\n` +
    `• תאריך החתימה: ${data.signDate || ""}`;
  
  const infoP = body.appendParagraph(infoText);
  applyRtl(infoP);

  // השמת החתימה במסגרת משמאל למטה
  if (sigBlob) {
    const sigTitle = body.appendParagraph("חתימת ההורה המלווה:");
    sigTitle.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    sigTitle.setBold(true);
    if (rtl) sigTitle.setTextDirection(rtl);

    // ב-Google Docs אי אפשר להוסיף תמונה ישירות אלא רק דרך הפסקה בצורה הזו:
    const img = body.appendImage(sigBlob);
    
    // כיווץ יחסי של התמונה כדי שתיכנס תמיד יפה בצד
    const maxWidth = 130;
    const width = img.getWidth() || maxWidth;
    const height = img.getHeight() || 60;
    img.setWidth(maxWidth);
    img.setHeight(maxWidth * (height / width));
  }

  doc.saveAndClose();
  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdfBlob;
}