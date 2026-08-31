function getParticipantsForEvent_(eventId) {
  const rows = sh_(APP.SHEETS.PARTICIPANTS).getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => ({
    id: r[0], personId: r[2], name: r[3], surname: r[4], club: r[5], role: r[6],
    email: r[7], refundLimitOverride: r[8], source: r[9], sourceFile: r[10], notes: r[11]
  }));
}

function participantHasRefundOverride_(participant) {
  return participant && participant.refundLimitOverride !== '' && participant.refundLimitOverride !== null && participant.refundLimitOverride !== undefined;
}

function saveParticipantRefundOverrides_(eventId, updates) {
  updates = updates || [];
  if (!eventId || !updates.length) return;

  const sheet = sh_(APP.SHEETS.PARTICIPANTS);
  const rows = sheet.getDataRange().getValues();
  const byId = {};
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(eventId)) byId[String(rows[i][0])] = i + 1;
  }

  updates.forEach(u => {
    const rowNumber = byId[String(u.id || '')];
    if (!rowNumber) return;
    if (u.reset === true) {
      sheet.getRange(rowNumber, 9).clearContent();
      return;
    }
    const value = Number(u.value);
    if (!isFinite(value) || value < 0) throw new Error('Il rimborso personale deve essere un importo uguale o superiore a zero.');
    sheet.getRange(rowNumber, 9).setValue(value);
  });
}

function getParticipantRefundLimitForBeneficiary_(eventId, event, beneficiary) {
  const key = normalize_(beneficiary);
  const participants = getParticipantsForEvent_(eventId);
  const participant = participants.find(p => normalize_((p.name || '') + ' ' + (p.surname || '')) === key) || null;
  if (participant && participantHasRefundOverride_(participant)) {
    return { value: Number(participant.refundLimitOverride || 0), source: 'PERSONALE', participant: participant };
  }
  const eventLimit = getEventRefundLimit_(eventId, event);
  return { value: Number(eventLimit.value || 0), source: eventLimit.source || 'EVENTO', participant: participant };
}

function findKnownPerson_(name, surname, cardNumber) {
  const rows = sh_(APP.SHEETS.PEOPLE).getDataRange().getValues();
  const card = String(cardNumber || '').trim();
  const key = samePersonKey_(name, surname);
  for (let i = 1; i < rows.length; i++) {
    if (card && String(rows[i][5] || '').trim() === card) {
      return { row:i+1, id:rows[i][0], name:rows[i][1], surname:rows[i][2], email:rows[i][3], club:rows[i][4], card:rows[i][5], type:rows[i][6] };
    }
  }
  for (let i = 1; i < rows.length; i++) {
    if (samePersonKey_(rows[i][1], rows[i][2]) === key) {
      return { row:i+1, id:rows[i][0], name:rows[i][1], surname:rows[i][2], email:rows[i][3], club:rows[i][4], card:rows[i][5], type:rows[i][6] };
    }
  }
  return null;
}

function findKnownEmail_(name, surname, cardNumber) {
  const person = findKnownPerson_(name, surname, cardNumber);
  return person && person.email ? String(person.email).trim() : '';
}

function saveKnownPerson(name, surname, club, email, cardNumber, type, notes) {
  const sheet = sh_(APP.SHEETS.PEOPLE);
  const person = findKnownPerson_(name, surname, cardNumber);
  const now = new Date();
  if (person) {
    sheet.getRange(person.row, 2, 1, 8).setValues([[
      name || person.name, surname || person.surname, email || person.email, club || person.club,
      cardNumber || person.card, type || person.type, now, notes || ''
    ]]);
    return person.id;
  }
  const id = 'P-' + Utilities.getUuid();
  sheet.appendRow([id, name || '', surname || '', email || '', club || '', cardNumber || '', type || 'ATLETA', now, notes || '']);
  return id;
}

function saveKnownEmail(name, surname, club, email) {
  if (!email) throw new Error('Email mancante.');
  return saveKnownPerson(name, surname, club, email, '', '', '');
}
