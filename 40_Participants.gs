function getParticipantsForEvent_(eventId) {
  const rows = sh_(APP.SHEETS.PARTICIPANTS).getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => ({
    id: r[0], name: r[3], surname: r[4], club: r[5], email: r[7], refundLimitOverride: r[9], notes: r[11]
  }));
}

function findKnownEmail_(name, surname) {
  const key = samePersonKey_(name, surname);
  const rows = sh_(APP.SHEETS.PEOPLE).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (samePersonKey_(rows[i][1], rows[i][2]) === key && rows[i][5]) return String(rows[i][5]).trim();
  }
  return '';
}

function saveKnownEmail(name, surname, club, email) {
  if (!email) throw new Error('Email mancante.');
  const sheet = sh_(APP.SHEETS.PEOPLE);
  const rows = sheet.getDataRange().getValues();
  const key = samePersonKey_(name, surname);
  for (let i = 1; i < rows.length; i++) {
    if (samePersonKey_(rows[i][1], rows[i][2]) === key) {
      sheet.getRange(i + 1, 4).setValue(club || rows[i][3]);
      sheet.getRange(i + 1, 6).setValue(email);
      sheet.getRange(i + 1, 9).setValue(new Date());
      return;
    }
  }
  sheet.appendRow(['P-' + Utilities.getUuid(), name, surname, club || '', '', email, '', '', new Date()]);
}
