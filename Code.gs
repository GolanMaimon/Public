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

    // המרת חתימה ל-Blob (ללא שמירת קובץ נפרד בדרייב)
    let sigLink = "";
    let sigBlob = null;
    if (data.signature && String(data.signature).startsWith("data:image")) {
      logMsg(runId, "signature detected");
      const base64 = data.signature.split(",")[1];
      const bytes = Utilities.base64Decode(base64);
      sigBlob = Utilities.newBlob(bytes, "image/png", "signature.png");
      logMsg(runId, "signature blob created for PDF only");
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
  const p = body.appendParagraph(`${rlm}${index}. ${text}${rlm}`);
  applyRtl(p);
  return p;
}



function buildPdfBlob(data, sigBlob) {
  const doc = DocumentApp.create("temp_pdf");
  const body = doc.getBody();
  
  // הצרת שוליים משמעותית כדי להכניס הכל לעמוד אחד מאורגן
  // נקטין את השלי העליון והתחתון אפילו יותר (מ-10 ל-5)
  body.setMarginTop(5).setMarginBottom(5).setMarginLeft(15).setMarginRight(15);
  
  // הגדרת פונט קטן יותר וצפיפות כדי לוודא שזה נכנס לדף אחד
  const style = {};
  style[DocumentApp.Attribute.FONT_SIZE] = 8;
  style[DocumentApp.Attribute.LINE_SPACING] = 0.8; // הוקטן מ-0.85
  body.setAttributes(style);

  const rtl = (DocumentApp.TextDirection && DocumentApp.TextDirection.RIGHT_TO_LEFT) || null;
  if (rtl) body.setTextDirection(rtl);

  const rlm = "\u200f";
  
  const title = body.appendParagraph(`${rlm}טופס אישור השתתפות במסיבת סיום שכבת כיתות ב'${rlm}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .setBold(true);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  if (rtl) title.setTextDirection(rtl);

  const intro = body.appendParagraph(
    `${rlm}הורים ותלמידים יקרים, לקראת מסיבת הסיום השנתית של שכבת כיתות ב\' (כיתות ב\'1 ו-ב\'2) ` +
    `מבית ספר "ערמונים", אשר תתקיים בתאריך 10/6/2026, בין השעות 16:00-19:00, במתחם בריכת כפר מעש, ` +
    `אנו מבקשים מכם לקרוא בעיון רב מסמך זה, לחתום עליו ולהחזירו לנציגי הוועד.${rlm}`
  ).setBold(true);
  applyRtl(intro);
  
  const separator = body.appendParagraph("________________________________________").setBold(true);
  applyRtl(separator);

  const items = [
    "לתשומת לבכם, השתתפותו/ה של התלמיד/ה באירוע מותנית בהמצאת טופס זה כשהוא מלא וחתום במלואו, וכן בליווי הורה כפי שיפורט להלן.",
    "\"הוועד\" או \"נציגי הוועד\" הם הורים בבית ספר \"ערמונים\", הפועלים בהתנדבות מלאה.",
    "מובהר בזאת, כי תפקידם של נציגי הוועד מוגבל אך ורק לתיאום הלוגיסטי והטכני המקדים של מסיבת הסיום. הוועד ו/או מי מהורי הוועד אינם משמשים כמפיקים, מפעילים, משגיחים, מצילים או אחראים על ביטחונם של הילדים המשתתפים במהלך האירוע עצמו.",
    "חלה חובה מוחלטת על נוכחות פיזית של הורה (או אפוטרופוס חוקי) מלווה עבור כל ילד וילדה לאורך כל שעות הפעילות במסיבת הסיום במתחם הבריכה.",
    "לא תותר כניסה או שהות של ילד או ילדה למתחם הבריכה ללא השגחה צמודה ורציפה של ההורה שלו/ה.",
    "ההורה המלווה נושא באחריות המלאה להשגחה על ילדו/ה, בדגש על שהות במים ובסביבת הבריכה.",
    "מובהר ומוסכם בזאת כי אין ולא תהיה כל אחריות במישרין ו/או בעקיפין להורי ועד כיתות ב' בבי״ס ערמונים לכל נזק (גוף, נפש או רכוש), פציעה, אובדן או פגיעה שיגרמו לילד/ה או לרכושו/ה במהלך האירוע, בדרך אליו או בחזרה ממנו.",
    "מוסכם בזאת, כי האחריות המלאה, הבלעדית והמוחלטת לשלומו, ביטחונו, בריאותו והתנהגותו של הילד/ה מוטלת על ההורים בלבד, ולא על נציגי הוועד.",
    "מובהר ומוסכם, כי בכפוף להשגחת ההורה, כל ילד/ה נושא/ת באחריות אישית על התנהגותו/ה במהלך האירוע, על כל המשתמע מכך, ובמנותק לחלוטין מפעילות הוועד.",
    "ההורים מצהירים בזאת כי בנם/בתם יודע/ת לשחות בצורה טובה ומשביעת רצון, ומורשה/ית להיכנס למים. ילד או ילדה שאינם יודעים לשחות בצורה טובה, מתחייבים לשהות אך ורק במים הרדודים ו/או להצטייד במצופים, הכל ובנוסף לאחריות ההורה המלווה.",
    "ההורים מתחייבים כי שוחחו עם הילד/ה והנחו אותו/ה להישמע באופן מוחלט להוראות הבטיחות של הנהלת הבריכה, המצילים ואנשי הצוות במקום, ויפקחו על כך באופן אישי במהלך האירוע.",
    "ההורים מצהירים כי הילד/ה כשיר/ה מבחינה בריאותית, גופנית ונפשית להשתתף באירוע ובפעילות במים, וכי אין מניעה רפואית כלשהי להשתתפותו/ה.",
    "מובהר ומוסכם בזאת, כי ההורים מוותרים ויתור סופי, מוחלט, מלא ובלתי חוזר על כל טענה, דרישה, דרישת פיצוי, קובלנה ו/או תביעה משפטית או אחרת, מכל סיבה שהיא, כנגד הורי ועד כיתות ב' 1+2 ו/או כנגד מי מטעמם.",
    "למען הסר ספק מובהר בזאת, ככל שתוגש תביעה או דרישה בניגוד לאמור במסמך זה, מתחייבים ההורים החתומים מטה לשאת בכל ההוצאות והנזקים שיגרמו לוועד או לחבריו בעקבות זאת."
  ];

  items.forEach((t, i) => {
    let p = addRtlNumbered(body, i + 1, t);
    // איפוס מפורש ל-Bold כדי שלא ירש את ההדגשה מהפסקאות הקודמות (Google Docs מוריש עיצוב מפסקה לפסקה)
    p.editAsText().setBold(false);

    // הדגשת המילים "לתשומת לבכם" אם קיימות בסעיף
    let textObj = p.editAsText();
    let textStr = textObj.getText();
    let boldIndex = textStr.indexOf("לתשומת לבכם");
    if (boldIndex !== -1) {
      textObj.setBold(boldIndex, boldIndex + "לתשומת לבכם".length - 1, true);
    }
  });

  const confirmText = `${rlm}אני החתום/ה מטה, הורה ו/או אפוטרופוס חוקי של התלמיד/ה, מאשר/ת בזאת כי קראתי בעיון את כל הסעיפים המופיעים במסמך זה, הבנתי את משמעותם המשפטית והכללית, ואני מסכים/ה להם ומתחייב/ת לפעול על פיהם במלואם ובאופן בלתי מסויג, לרבות הצהרתי כי בדקתי את השטח ומצאתיו תקין, והתחייבותי לנכוח באופן אישי במסיבת הסיום ולהשגיח על ילדי/בתי.${rlm}`;
  const confirm = body.appendParagraph(confirmText).setBold(true);
  applyRtl(confirm);

  const fields = [
    { label: "• שם מלא של התלמיד/ה:", val: (data.childName || "") + "\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0 כיתה: " + (data.childClass || "") },
    { label: "• תעודת זהות של התלמיד/ה:", val: data.childId },
    { label: "• שם מלא של ההורה המלווה שינכח באירוע:", val: data.parentName },
    { label: "• תעודת זהות של ההורה המלווה:", val: data.parentId },
    { label: "• כתובת מגורים:", val: data.address },
    { label: "• טלפון נייד של ההורה המלווה:", val: data.phone },
    { label: "• טלפון נוסף לשעת חירום:", val: data.phone2 }
  ];

  // יצירת טבלה ללא ריווחים ליישור מדויק אחד מתחת לשני
  const table = body.appendTable();
  table.setBorderWidth(0);

  fields.forEach(f => {
    let tr = table.appendTableRow();
    
    // תא התשובה (השמאלי - מתווסף ראשון בחוקיות ltr)
    let valCell = tr.appendTableCell();
    
    // בנייה בטוחה של הפסקה בתא הראשון
    let pVal = valCell.getChild(0).asParagraph();
    // רק את הטקסט, והשאר רווחים שמאלה (ללא הנקודה)
    let valStr = f.val ? String(f.val).trim() : "";
    let baseStr = valStr + "\u00A0"; 
    while (baseStr.length < 47) { 
      baseStr += "\u00A0";
    }

    // הוספת הטקסט לפסקה: תגית rlm עוטפת הכל כדי לשמור על עברית מלאה ורציפה
    pVal.setText(`${rlm}${baseStr}${rlm}`);
    applyRtl(pVal);
    // יישור לימין!
    pVal.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    pVal.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.0).setBold(true);
    
    // קביעת הקו התחתון לכל הטקסט (כולל 
    // קביעת הקו התחתון לכל הטקסט (כולל הנקודה והרווחים) כדי שיראה כקו אחד
    pVal.editAsText().setUnderline(0, pVal.getText().length - 1, true);

    // תא השם (הימני - מתווסף שני)
    let labelCell = tr.appendTableCell();
    labelCell.setWidth(210); // רוחב קבוע כדי שכולם יתחילו באותה נקודה
    
    // בנייה בטוחה של הפסקה בתא השני
    let pLabel = labelCell.getChild(0).asParagraph();
    // כדי למקם את הנקודתיים במדויק משמאל לשדה בממשק של גוגל, עדיף פשוט לכתוב אותן באנגלית בסוף המחרוזת מבחינה לוגית
    pLabel.setText(`${rlm}${f.label.replace(":", "")} :${rlm}`);
    applyRtl(pLabel);
    pLabel.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    pLabel.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.0).setBold(true);
  });

  // טיפול בתאריך והצגתו בפורמט DD/MM/YYYY
  let signDateStr = data.signDate;
  if (signDateStr && signDateStr.includes("-")) {
    let parts = signDateStr.split("-");
    if (parts.length === 3) {
      signDateStr = `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  } else {
    signDateStr = Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yyyy");
  }
  
  // יצירת טבלת מכולה (ללא גבולות) כדי לשים את התאריך והחתימה באותו קו גובה בדיוק
  const footerTable = body.appendTable();
  footerTable.setBorderWidth(0); 
  const footerRow = footerTable.appendTableRow();
  
  // תא שמאלי: חתימה
  const sigCellOuter = footerRow.appendTableCell();
  sigCellOuter.setWidth(250); // מגביל את התא לרוחב המסגרת, הוגדל ל-250 לפי הצורך

  // תא ימני: תאריך
  const dateCellOuter = footerRow.appendTableCell();

  if (sigBlob) {
    // השמת כותרת החתימה במרכז הקו העליון של המסגרת
    let sigTitle = sigCellOuter.getChild(0).asParagraph();
    sigTitle.setText(`${rlm}חתימת ההורה המלווה:${rlm}`);
    applyRtl(sigTitle);
    sigTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER); // מרכז!
    sigTitle.setBold(true).setSpacingBefore(0).setSpacingAfter(0);

    try {
      // מסגרת החתימה (טבלה פנימית קטנה עם גבולות)
      const sigTableInner = sigCellOuter.appendTable();
      sigTableInner.setBorderWidth(1); 
      let sigRowInner = sigTableInner.appendTableRow();
      let sigCellInner = sigRowInner.appendTableCell();
      
      // איפוס שוליים למסגרת עם מעט ריווח לאיור
      sigCellInner.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(0).setPaddingRight(0);
      
      let sigImgP = sigCellInner.getChild(0).asParagraph();
      sigImgP.setAlignment(DocumentApp.HorizontalAlignment.CENTER); // גם התמונה תמורכז
      sigImgP.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.0);

      const img = sigImgP.appendInlineImage(sigBlob);
      const targetWidth = 250; 
      const width = img.getWidth() || targetWidth;
      const height = img.getHeight() || 60;
      img.setWidth(targetWidth);
      img.setHeight(targetWidth * (height / width));
    } catch (err) {
      Logger.log("Signature append error: " + err);
    }
  } else {
    sigCellOuter.getChild(0).asParagraph().setText("");
  }

  // השמת התאריך בתא הימני - מופיע באותו קו גובה אופקי של הכותרת השמאלית
  let dateP = dateCellOuter.getChild(0).asParagraph();
  dateP.setText(`${rlm}תאריך: ${signDateStr}${rlm}`);
  applyRtl(dateP);
  dateP.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  dateP.setBold(true).setSpacingBefore(0).setSpacingAfter(0);

  // הסרת פסקאות ריקות שעשויות לגרום לעמוד נוסף למטה ולמעלה
  // 1. הפסקה הריקה הראשונה שנוצרת אוטומטית יחד עם מסמך חדש
  if (body.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH && body.getChild(0).asParagraph().getText().trim() === "") {
    body.removeChild(body.getChild(0));
  }
  
  // 2. הפסקה האחרונה (מנוע גוגל דוקס מחייב פסקת טקסט אחרי טבלה, לא ניתן למחוק אותה אז נקטין אותה למינימום מוחלט)
  let finalChild = body.getChild(body.getNumChildren() - 1);
  if (finalChild.getType() === DocumentApp.ElementType.PARAGRAPH && finalChild.asParagraph().getText().trim() === "") {
    finalChild.asParagraph().setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(0.06);
    finalChild.asParagraph().editAsText().setFontSize(1);
  }

  doc.saveAndClose();
  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdfBlob;
}