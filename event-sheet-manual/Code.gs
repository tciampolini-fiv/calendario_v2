/*
 * SCHEDA EVENTO - MODALITA MANUALE TEMPORANEA
 * 14/09/2026
 *
 * Questo script e' autonomo e va usato DIRETTAMENTE nel progetto Apps Script
 * collegato alla singola Scheda Evento, anche se vecchia.
 *
 * Obiettivo:
 * - lasciare Spese, Partecipanti e Documenti completamente manuali/invariati;
 * - rendere il foglio "Attività" libero da protezioni, gruppi, convalide e
 *   formattazioni automatiche che possono impedire o rallentare il lavoro;
 * - non avere nessun onEdit, autosync, riordino o ricostruzione automatica.
 *
 * Dopo aver sostituito il vecchio codice della Scheda Evento con questo file:
 * 1. salvare;
 * 2. ricaricare il Google Sheet;
 * 3. menu "Scheda Evento" -> "Sblocca Attività per lavoro manuale";
 * 4. autorizzare lo script, se richiesto.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Scheda Evento')
    .addItem('Sblocca Attività per lavoro manuale', 'sbloccaAttivitaManuale')
    .addItem('Rimuovi trigger automatici', 'rimuoviTriggerAutomatici')
    .addToUi();
}

/**
 * Rende il SOLO foglio "Attività" completamente manuale.
 * Non cancella obiettivi, task, scadenze, note o altri dati inseriti.
 */
function sbloccaAttivitaManuale() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Attività') || ss.getSheetByName('Attivita');

  if (!sh) {
    SpreadsheetApp.getUi().alert('Foglio "Attività" non trovato.');
    return;
  }

  let protezioniRimosse = 0;
  let gruppiRimossi = 0;

  // 1) Rimuove tutte le protezioni SOLO dal foglio Attività.
  [
    SpreadsheetApp.ProtectionType.SHEET,
    SpreadsheetApp.ProtectionType.RANGE
  ].forEach(function(type) {
    sh.getProtections(type).forEach(function(p) {
      try {
        p.remove();
        protezioniRimosse++;
      } catch (err) {
        console.log('Protezione non rimossa: ' + err.message);
      }
    });
  });

  // 2) Rimuove i gruppi righe: nessun ordine/struttura bloccata dei task.
  try {
    for (let r = 1; r <= sh.getMaxRows(); r++) {
      const group = sh.getRowGroup(r, 1);
      if (group) {
        group.remove();
        gruppiRimossi++;
      }
    }
  } catch (err) {
    console.log('Rimozione gruppi: ' + err.message);
  }

  // 3) Mostra tutte le righe e colonne eventualmente nascoste.
  try {
    sh.showRows(1, sh.getMaxRows());
  } catch (err) {}
  try {
    sh.showColumns(1, sh.getMaxColumns());
  } catch (err) {}

  // 4) Rimuove checkbox, dropdown e altre convalide che possono bloccare
  //    l'inserimento manuale. I valori gia' presenti NON vengono cancellati.
  const rows = Math.max(1, sh.getMaxRows());
  const cols = Math.max(1, sh.getMaxColumns());
  sh.getRange(1, 1, rows, cols).clearDataValidations();

  // 5) Elimina la formattazione condizionale automatica.
  try {
    sh.setConditionalFormatRules([]);
  } catch (err) {
    console.log('Formattazione condizionale: ' + err.message);
  }

  // 6) Neutralizza la formattazione visibile senza cancellare dati,
  //    formule o formati data/numero.
  const usedRows = Math.max(1, sh.getLastRow());
  const usedCols = Math.max(1, sh.getLastColumn());
  const used = sh.getRange(1, 1, usedRows, usedCols);
  used
    .setBackground(null)
    .setFontWeight('normal')
    .setFontStyle('normal')
    .setFontColor(null);

  // 7) Elimina soltanto i vecchi controlli grafici "+" della versione
  //    automatica. Non tocca obiettivi e task reali.
  if (usedRows >= 2) {
    const scanCols = Math.min(8, sh.getMaxColumns());
    const range = sh.getRange(2, 1, usedRows - 1, scanCols);
    const values = range.getValues();
    let changed = false;

    values.forEach(function(row) {
      if (scanCols >= 3) {
        const c = String(row[2] || '').trim();
        if (c === '+' || c === '＋') {
          row[2] = '';
          changed = true;
        }
      }

      const a = String(row[0] || '').trim();
      const marker = scanCols >= 8 ? String(row[7] || '').trim().toUpperCase() : '';
      if (/^[+＋]\s*AGGIUNGI OBIETTIVO/i.test(a) || marker === 'AGGIUNGI_OBIETTIVO') {
        for (let c = 0; c < scanCols; c++) row[c] = '';
        changed = true;
      }
    });

    if (changed) range.setValues(values);
  }

  // 8) Elimina gli eventuali trigger installabili del progetto della Scheda.
  //    Nessun trigger automatico deve continuare a modificare Attività.
  const triggerRimossi = rimuoviTriggerAutomatici_(false);

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Attività pronta per il lavoro manuale',
    'Protezioni rimosse: ' + protezioniRimosse + '\n' +
    'Gruppi righe rimossi: ' + gruppiRimossi + '\n' +
    'Trigger automatici rimossi: ' + triggerRimossi + '\n\n' +
    'Ora puoi inserire, modificare, spostare o cancellare liberamente obiettivi e task.\n' +
    'Le altre schede non sono state modificate.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Comando manuale separato per eliminare gli eventuali trigger installabili
 * rimasti nel progetto Apps Script della Scheda Evento.
 */
function rimuoviTriggerAutomatici() {
  const count = rimuoviTriggerAutomatici_(true);
  return count;
}

function rimuoviTriggerAutomatici_(showAlert) {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    try {
      ScriptApp.deleteTrigger(trigger);
      count++;
    } catch (err) {
      console.log('Trigger non rimosso: ' + err.message);
    }
  });

  if (showAlert) {
    SpreadsheetApp.getUi().alert(
      'Trigger automatici',
      'Trigger rimossi dal progetto della Scheda Evento: ' + count,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }

  return count;
}
