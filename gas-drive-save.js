// ================================================================
// SpeedLabor – Live-Sync + PDF-Speicherung in Google Drive
// Google Apps Script Web App  (VOLLSTÄNDIGE VERSION)
//
// EINMALIGE EINRICHTUNG (ca. 5 Minuten):
// 1. Öffne https://script.google.com
// 2. Klicke "+ Neues Projekt", nenne es "SpeedLabor"
// 3. Lösche den Beispielcode, füge DIESEN Code ein
// 4. Ändere FOLDER_ID zur ID deines Drive-Ordners
//    (Drive-Ordner öffnen -> URL: .../folders/HIER_IST_DIE_ID)
// 5. Klicke oben "Bereitstellen" -> "Neue Bereitstellung"
//    - Typ: Web App
//    - Ausführen als: Ich
//    - Zugriff: Jeder
// 6. Klicke "Bereitstellen" -> Berechtigungen erlauben
// 7. Kopiere die Web-App-URL
// 8. Trage diese URL in reparaturauftrag.html UND monitor.html
//    bei: const GAS_URL = 'HIER_GAS_URL_EINTRAGEN';
// ================================================================

const FOLDER_ID = '1r_vsZ64V-35M1Il1bpTDmFqj14oBTSb1'; // <- deine Drive-Ordner-ID

// ---------------------------------------------------------------
// Hilfsfunktion: Spreadsheet fuer Draft-Daten holen/erstellen
// ---------------------------------------------------------------
function getSheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('ssId');
  if (!ssId) {
    const ss = SpreadsheetApp.create('SpeedLabor_Live_Draft');
    ssId = ss.getId();
    props.setProperty('ssId', ssId);
    const sheet = ss.getSheets()[0];
    sheet.setName('draft');
    sheet.getRange('A1:D1').setValues([['', '', 'empty', '']]);
  }
  const ss = SpreadsheetApp.openById(ssId);
  return ss.getSheetByName('draft') || ss.getSheets()[0];
}

// ---------------------------------------------------------------
// POST: syncDraft / clearDraft / confirmPrint / savePDF
// ---------------------------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    switch (data.action) {

      case 'syncDraft': {
        const fd = data.formData || {};
        const sig = fd.signature || '';
        const fieldsOnly = Object.assign({}, fd);
        delete fieldsOnly.signature;
        // Assign order number if not yet set for this draft
        const existing = sheet.getRange('A1:E1').getValues()[0];
        const prevStatus = existing[2] || 'empty';
        // Don't overwrite if already printed – let tablet detect it
        if (prevStatus === 'printed') {
          return jsonResponse({ success: true, orderNumber: existing[4], status: 'printed' });
        }
        let orderNumber = (prevStatus === 'empty') ? null : existing[4];
        if (!orderNumber) {
          const props = PropertiesService.getScriptProperties();
          const counter = parseInt(props.getProperty('orderCounter') || '1000', 10) + 1;
          props.setProperty('orderCounter', String(counter));
          orderNumber = counter;
        }
        sheet.getRange('A1:E1').setValues([[
          JSON.stringify(fieldsOnly),
          sig,
          'pending',
          new Date().toISOString(),
          orderNumber
        ]]);
        return jsonResponse({ success: true, orderNumber });
      }

      case 'clearDraft':
        sheet.getRange('A1:E1').setValues([['', '', 'empty', '', '']]);
        return jsonResponse({ success: true });

      case 'confirmPrint':
        sheet.getRange('C1').setValue('printed');
        return jsonResponse({ success: true });

      case 'savePDF': {
        const pdfBytes = Utilities.base64Decode(data.pdfBase64);
        const blob = Utilities.newBlob(pdfBytes, 'application/pdf', data.filename);
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const file = folder.createFile(blob);
        return jsonResponse({ success: true, fileUrl: file.getUrl() });
      }

      default:
        return jsonResponse({ success: false, error: 'Unknown action: ' + data.action });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ---------------------------------------------------------------
// GET: getDraft / Status-Seite
// ---------------------------------------------------------------
function doGet(e) {
  if (e.parameter && e.parameter.action === 'getDraft') {
    try {
      const sheet = getSheet();
      const row = sheet.getRange('A1:E1').getValues()[0];
      const fieldsJson   = row[0];
      const signature    = row[1];
      const status       = row[2] || 'empty';
      const time         = row[3] ? String(row[3]) : '';
      const orderNumber  = row[4] || '';

      if (!fieldsJson) {
        return jsonResponse({ formData: null, status: 'empty', time: '', orderNumber: '' });
      }
      const formData = JSON.parse(fieldsJson);
      formData.signature = signature;
      return jsonResponse({ formData, status, time, orderNumber });
    } catch (err) {
      return jsonResponse({ error: err.message });
    }
  }

  return HtmlService.createHtmlOutput(
    '<h2 style="font-family:sans-serif">SpeedLabor GAS OK</h2><p>Web App laeuft.</p>'
  );
}

// ---------------------------------------------------------------
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
