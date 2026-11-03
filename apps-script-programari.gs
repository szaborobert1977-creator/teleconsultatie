// Se lipește în Google Sheets: Extensii > Apps Script
const DOCTOR = 'szaborobert1977@gmail.com';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Programari');
  if (!sh) {
    sh = ss.insertSheet('Programari');
    sh.appendRow(['Înregistrat', 'Data', 'Ora', 'Nume', 'Telefon', 'Email', 'Motiv']);
  }
  sh.getRange('B:C').setNumberFormat('@'); // data și ora rămân text
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Returnează intervalele deja ocupate
function doGet() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  const booked = n > 0 ? sh.getRange(2, 2, n, 2).getValues().map(r => r[0] + ' ' + r[1]) : [];
  return json_({ booked: booked });
}

// Primește o programare nouă
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents);
    const sh = sheet_();
    const n = sh.getLastRow() - 1;
    const taken = n > 0 ? sh.getRange(2, 2, n, 2).getValues().some(r => r[0] === p.data && r[1] === p.ora) : false;
    if (taken) return json_({ ok: false, error: 'taken' });

    sh.appendRow([new Date(), p.data, p.ora, p.nume, p.telefon, p.email || '', p.motiv || '']);

    MailApp.sendEmail(DOCTOR, 'Programare nouă – ' + p.eticheta,
      'Programare: ' + p.eticheta + '\nNume: ' + p.nume + '\nTelefon: ' + p.telefon +
      '\nE-mail: ' + (p.email || '-') + '\nMotiv: ' + (p.motiv || '-'));

    if (p.email) {
      try {
        MailApp.sendEmail(p.email, 'Cererea ta de programare',
          'Bună, ' + p.nume + '!\n\nAm primit cererea ta pentru ' + p.eticheta +
          '.\nVei fi contactat pentru confirmare.');
      } catch (err) {}
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
