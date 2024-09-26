/**
 * This App Script update the calendar sheet with the planning tables for the current week and the following ones.
 * On a new week, it removes the past one and adds a empty new one at the end.
 * Named ranges are used to parameter all.
 */

const APP_TITLE = "Agenda Atelier Nuances";

const PUBLIC_CALENDAR_SHEET_ID = "1_0Mh-E4UW4-eC-Y6oMr3VwMoKZSNE-i3woWh5OpXDmA";

const CALENDAR_SHEET_NAME = "Calendrier Céramistes";
const HOMEPAGE_SHEET_NAME = "INFOS";
const PEOPLE_SHEET_NAME = "Inscrits";
const OPENINGS_SHEET_NAME = "Ouvertures";
const CLOSED_SHEET_NAME = "Fermetures";
const CATEGORIES_SHEET_NAME = "Catégories";
const PARAMETERS_SHEET_NAME = "Paramètres";
const SAVE_SHEET_NAME = "SaveData";

/** Should the data in the calendar be conserved. If false, it will be generated empty. */
const KEEP_CALENDAR_DATA = true;
/** Should people past days counts be updated by adding the days removed and prior to today. */
const UPDATE_PEOPLE_PAST_DAYS = true;

const TYPES_OF_PEOPLE = ['Tourneurs', 'Modeleurs'];

const SELF_DAYS_HEADER = "Zone Libre";

const PEOPLE_HEADER_NB_ROWS = 3;


/**
 * Regenerate the calendar
 */
function generateCalendar() {
  let publicSpreadsheet = SpreadsheetApp.openById(PUBLIC_CALENDAR_SHEET_ID);
  if (!publicSpreadsheet) {
    err(`Impossible d'ouvrir le calendrier public.`);
    return;
  }

  let calendarSheet = publicSpreadsheet.getSheetByName(CALENDAR_SHEET_NAME);

  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let saveSheet = activeSpreadsheet.getSheetByName(SAVE_SHEET_NAME);
  let p = new GenerateParameters(activeSpreadsheet, publicSpreadsheet);

  let startingWeekRow = 2;
  let weekCol = 1;

  // -- Set wip message
  let messageRange = calendarSheet.getRange("F1");
  messageRange.setValue("MISE À JOUR EN COURS, PATIENTEZ");
  let errorRange = activeSpreadsheet.getSheetByName(HOMEPAGE_SHEET_NAME).getRange(3, 2);
  errorRange.setValue("MISE À JOUR EN COURS, ne PAS fermer la page");

  // -- Update people list
  updatePublicPeople(activeSpreadsheet, publicSpreadsheet, p.peopleNames);

  info("Mise à jour démarrée, ne pas fermer la page.");

  // -- Copy people data from currently displayed weeks
  let calendarRange = calendarSheet.getRange(
    startingWeekRow + 1, weekCol,
    calendarSheet.getMaxRows() - startingWeekRow, calendarSheet.getMaxColumns() - weekCol + 1
  );

  let savedValues = calendarRange.getDisplayValues();
  let saveRange = saveSheet.getRange(1, 1, calendarRange.getNumRows(), calendarRange.getNumColumns())
    .clear()
    .clearDataValidations()
    .setNumberFormat("@");

  // Copy to save sheet
  saveRange.setValues(savedValues);

  log(`Calendar saved.`);

  // -- Try to update the sheet and if there is an issue, copy back the saved sheet
  try {
    // -- Generate saved values map
    log(`Generate saved values map`);
    let [savedDaysMap, savedSelfDaysMap] = getDaysMap(savedValues);

    // -- Create header
    const headerNumRows = addHeaderToCalendar(p, calendarSheet, startingWeekRow, weekCol);
    log(`Added header.`);

    let weekRow = startingWeekRow + headerNumRows;

    /** @type {GoogleAppsScript.Spreadsheet.Range[]} */
    let weeksSubheaderRanges = [];
    /** @type {GoogleAppsScript.Spreadsheet.Range[]} */
    let weeksSeparatorRanges = [];

    /** @type {String[][]} */
    let newValues = [];
    /** @type {GoogleAppsScript.Spreadsheet.TextStyle[][]} */
    let newTextStyles = [];
    /** @type {String[][]} */
    let newHorizontalAlignments = [];
    /** @type {String[][]} */
    let newVerticalAlignments = [];
    /** @type {String[][]} */
    let newBackgrounds = [];
    /** @type {GoogleAppsScript.Spreadsheet.DataValidation[][]} */
    let newDataValidations = [];

    // -- Add weeks calendars, starting from current
    let year = p.today.getFullYear();
    let weekNo = getWeekNumber(p.today);
    const nbCols = calendarSheet.getMaxColumns() - weekCol + 1;
    for (let weekIdx = 0; weekIdx < p.weeksToDisplay; weekIdx++) {

      // Make sure weekNo is valid
      if (weekNo + weekIdx > 52) {
        year++;
        weekNo -= 52;
      }

      const prevRow = weekRow;

      const dateOfWeek = getDateOfWeek(weekNo + weekIdx, year);

      // -- Week subheader
      newValues.push(createRowValues(nbCols, "Semaine " + (weekNo + weekIdx) + " - " + year));
      newTextStyles.push(createRowValues(nbCols, p.subheaderTextStyle));
      newHorizontalAlignments.push(createRowValues(nbCols, "center"));
      newVerticalAlignments.push(createRowValues(nbCols, "middle"));
      newBackgrounds.push(createRowValues(nbCols, p.subheaderBackground));
      newDataValidations.push(createRowValues(nbCols));
      weeksSubheaderRanges.push(calendarSheet.getRange(weekRow, weekCol, 1, nbCols));

      weekRow += 1;

      // -- Regular openings
      for (let openingTime of p.openingTimes) {
        let begin = openingTime.getBeginInWeek(dateOfWeek);
        let end = openingTime.getEndInWeek(dateOfWeek);

        // Don't add the slot if closed
        if (p.isClosed(begin, end)) {
          log(`Not opened between ${begin} and ${end}.`);
          continue;
        }

        if (KEEP_CALENDAR_DATA)
          newValues.push(createOpeningRow(nbCols, openingTime, begin, end, p, savedDaysMap));
        else
          newValues.push(createOpeningRow(nbCols, openingTime, begin, end, p));
        newTextStyles.push(createRowValues(nbCols, [null, p.dayTextStyle, p.hoursTextStyle]));
        newHorizontalAlignments.push(createRowValues(nbCols, [null, "center", "center"]));
        newVerticalAlignments.push(createRowValues(nbCols, "middle", true));
        newBackgrounds.push(createRowValues(nbCols, [null, openingTime.dayColor, openingTime.hourColor]));
        newDataValidations.push(createRowValues(nbCols, [null, null, null, p.peopleRule], true));

        weekRow += 1;
      }

      // -- Self-openings
      let separatorInserted = false;
      for (let openingTime of p.selfopeningTimes) {
        let begin = openingTime.getBeginInWeek(dateOfWeek);
        let end = openingTime.getEndInWeek(dateOfWeek);

        // Don't add the slot if closed
        if (p.isClosed(begin, end)) {
          log(`Not opened between ${begin} and ${end}.`);
          continue;
        }

        // Separator
        if (!separatorInserted) {
          separatorInserted = true;

          newValues.push(createRowValues(nbCols, SELF_DAYS_HEADER));
          newTextStyles.push(createRowValues(nbCols, p.separatorTextStyle));
          newHorizontalAlignments.push(createRowValues(nbCols, "center"));
          newVerticalAlignments.push(createRowValues(nbCols, "middle"));
          newBackgrounds.push(createRowValues(nbCols, p.separatorBackground));
          newDataValidations.push(createRowValues(nbCols));
          weeksSeparatorRanges.push(calendarSheet.getRange(weekRow, weekCol, 1, nbCols));

          weekRow += 1;
        }

        if (KEEP_CALENDAR_DATA)
          newValues.push(createOpeningRow(nbCols, openingTime, begin, end, p, savedSelfDaysMap));
        else
          newValues.push(createOpeningRow(nbCols, openingTime, begin, end, p));
        newTextStyles.push(createRowValues(nbCols, [null, p.dayTextStyle, p.hoursTextStyle]));
        newHorizontalAlignments.push(createRowValues(nbCols, [null, "center", "center"]));
        newVerticalAlignments.push(createRowValues(nbCols, "middle", true));
        newBackgrounds.push(createRowValues(nbCols, [null, openingTime.dayColor, openingTime.hourColor]));
        newDataValidations.push(createRowValues(nbCols, [null, null, null, p.peopleRule], true));

        weekRow += 1;
      }

      log(`Added ${weekRow - prevRow} row for week ${weekNo + weekIdx} of ${year}.`);
    }

    log(`${weekRow - calendarRange.getRow()} rows filled.`);

    for (let i = weekRow - calendarRange.getRow(); i < calendarRange.getNumRows(); i++) {
      newValues.push(Array(nbCols));
      newTextStyles.push(Array(nbCols));
      newHorizontalAlignments.push(Array(nbCols));
      newVerticalAlignments.push(Array(nbCols));
      newBackgrounds.push(Array(nbCols));
      newDataValidations.push(Array(nbCols));
    }

    log(`Completed with ${calendarRange.getNumRows() - (weekRow - calendarRange.getRow())} empty rows for a total of ${calendarRange.getNumRows()} rows.`);

    if (newValues.length != calendarRange.getNumRows()) {
      throw new RangeError(`Inconsitent number of rows given as newValues (${newValues.length} while expecting ${calendarRange.getNumRows()}).`);
    }
    else if (newTextStyles.length != calendarRange.getNumRows()) {
      throw new RangeError(`Inconsitent number of rows given as newTextStyles (${newTextStyles.length} while expecting ${calendarRange.getNumRows()}).`);
    }
    else if (newHorizontalAlignments.length != calendarRange.getNumRows()) {
      throw new RangeError(`Inconsitent number of rows given as newHorizontalAlignments (${newHorizontalAlignments.length} while expecting ${calendarRange.getNumRows()}).`);
    }
    else if (newVerticalAlignments.length != calendarRange.getNumRows()) {
      throw new RangeError(`Inconsitent number of rows given as newVerticalAlignments (${newVerticalAlignments.length} while expecting ${calendarRange.getNumRows()}).`);
    }
    else if (newBackgrounds.length != calendarRange.getNumRows()) {
      throw new RangeError(`Inconsitent number of rows given as newBackgrounds (${newBackgrounds.length} while expecting ${calendarRange.getNumRows()}).`);
    }
    else if (newDataValidations.length != calendarRange.getNumRows()) {
      throw new RangeError(`Inconsitent number of rows given as newDataValidations (${newDataValidations.length} while expecting ${calendarRange.getNumRows()}).`);
    }

    // -- Clear
    calendarRange.clear()
      .clearDataValidations()
      .setNumberFormat("@");
    log(`Cleared the public calendar.`);

    calendarSheet.setRowHeights(1, calendarSheet.getMaxRows(), 21);

    log(`Merge weeks subheader and separators.`);
    for (let weekSubheaderRange of weeksSubheaderRanges) {
      weekSubheaderRange.mergeAcross()
        .setBorder(true, true, true, true, false, false, p.borderColor, null);
    }
    for (let weekSeparatorRange of weeksSeparatorRanges) {
      weekSeparatorRange.mergeAcross()
        .setBorder(true, true, true, true, false, false, p.borderColor, null);
    }

    // -- Set values
    log(`Setting calendarRange newValues.`);
    calendarRange.setValues(newValues);
    log(`Setting calendarRange newTextStyles.`);
    calendarRange.setTextStyles(newTextStyles);
    log(`Setting calendarRange newHorizontalAlignments.`);
    calendarRange.setHorizontalAlignments(newHorizontalAlignments);
    log(`Setting calendarRange newVerticalAlignments.`);
    calendarRange.setVerticalAlignments(newVerticalAlignments);
    log(`Setting calendarRange newBackgrounds.`);
    calendarRange.setBackgrounds(newBackgrounds);
    log(`Setting calendarRange newDataValidations.`);
    calendarRange.setDataValidations(newDataValidations);

    log(`calendarRange filled.`);

    // -- Update people future days formula
    if (true) {
      log(`Update people future days formula.`);

      let ceramistsFutureDays = [];
      let ceramistsSelfFutureDays = [];
      let modelersFutureDays = [];
      let modelersSelfFutureDays = [];

      let lastColName = columnToLetter(CALENDAR.SLOT + p.slotsNames.length);
      let selectedCols = Array.from(Array(p.ceramistsSlotsName.length), (_, i) => i + CALENDAR.SLOT).join(";");
      let filter = `FILTER(CHOOSECOLS('${CALENDAR_SHEET_NAME}'!$A:$${lastColName}; 1; ${selectedCols}); '${CALENDAR_SHEET_NAME}'!$A:$A = "${OPENING_TYPE.REGULAR}")`;
      let filterSelf = `FILTER(CHOOSECOLS('${CALENDAR_SHEET_NAME}'!$A:$${lastColName}; 1; ${selectedCols}); '${CALENDAR_SHEET_NAME}'!$A:$A = "${OPENING_TYPE.SELF}")`;
      for (let row = 0; row < p.ceramistsFutureDaysActiveRange.getNumRows(); row++) {
        ceramistsFutureDays.push([`=COUNTIF(${filter}; $A${PEOPLE_HEADER_NB_ROWS + row + 1})`]);
        ceramistsSelfFutureDays.push([`=COUNTIF(${filterSelf}; $A${PEOPLE_HEADER_NB_ROWS + row + 1})`]);
      }

      selectedCols = Array.from(Array(p.modelersSlotsName.length), (_, i) => i + CALENDAR.SLOT + p.ceramistsSlotsName.length).join(";");
      filter = `FILTER(CHOOSECOLS('${CALENDAR_SHEET_NAME}'!$A:$${lastColName}; 1; ${selectedCols}); '${CALENDAR_SHEET_NAME}'!$A:$A = "${OPENING_TYPE.REGULAR}")`;
      filterSelf = `FILTER(CHOOSECOLS('${CALENDAR_SHEET_NAME}'!$A:$${lastColName}; 1; ${selectedCols}); '${CALENDAR_SHEET_NAME}'!$A:$A = "${OPENING_TYPE.SELF}")`;
      for (let row = 0; row < p.modelersFutureDaysActiveRange.getNumRows(); row++) {
        modelersFutureDays.push([`=COUNTIF(${filter}; $A${PEOPLE_HEADER_NB_ROWS + row + 1})`]);
        modelersSelfFutureDays.push([`=COUNTIF(${filterSelf}; $A${PEOPLE_HEADER_NB_ROWS + row + 1})`]);
      }

      p.ceramistsFutureDaysActiveRange.setValues(ceramistsFutureDays);
      p.ceramistsSelfFutureDaysActiveRange.setValues(ceramistsSelfFutureDays);
      p.modelersFutureDaysActiveRange.setValues(modelersFutureDays);
      p.modelersSelfFutureDaysActiveRange.setValues(modelersSelfFutureDays);
    }

    // -- Update people past days counts
    if (UPDATE_PEOPLE_PAST_DAYS) {
      log(`Incrementing people past days.`);
      for (const [[b, e], r] of savedDaysMap) {
        if (e.getTime() <= p.today.getTime()) {
          for (let i = CALENDAR.SLOT; i < r.length; i++) {
            p.addPastDay(formatName(r[i]), i - CALENDAR.SLOT, false);
          }
        }
      }
      for (const [[b, e], r] of savedSelfDaysMap) {
        if (e.getTime() <= p.today.getTime()) {
          for (let i = CALENDAR.SLOT; i < r.length; i++) {
            p.addPastDay(formatName(r[i]), i - CALENDAR.SLOT, true);
          }
        }
      }

      p.ceramistsPastDaysActiveRange.setValues(p.ceramistsPastDays);
      p.ceramistsSelfPastDaysActiveRange.setValues(p.ceramistsSelfPastDays);
      p.modelersPastDaysActiveRange.setValues(p.modelersPastDays);
      p.modelersSelfPastDaysActiveRange.setValues(p.modelersSelfPastDays);
      log(`People past days updated.`);
    }

    errorRange.clearContent();
    messageRange.clearContent();
  } catch (e) {
    err(`Erreur pendant l'insertion des valeurs, sauvegarde restaurée.`, e);
    calendarRange.setValues(savedValues);
    errorRange.setValue(`Erreur prevenir Grégoire, ne rien toucher.`);
    messageRange.setValue(`Erreur prevenir Grégoire ou Karo, ne rien toucher.`);
  }

  // -- Set conditional format rules
  // Free slots
  let rules = [];
  let rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(p.freeSlotCell.getDisplayValue())
    .setFontColor(p.freeSlotCell.getFontColorObject().asRgbColor().asHexString())
    .setBackground(p.freeSlotCell.getBackground())
    .setRanges([calendarSheet.getRange(startingWeekRow, weekCol + CALENDAR.SLOT, calendarRange.getNumRows() - startingWeekRow + 1, p.slotsNames.length)])
    .build();
  rules.push(rule);

  // Unavailable slots
  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(p.unavailableSlotCell.getDisplayValue())
    .setFontColor(p.unavailableSlotCell.getFontColorObject().asRgbColor().asHexString())
    .setBackground(p.unavailableSlotCell.getBackground())
    .setRanges([calendarSheet.getRange(startingWeekRow, weekCol + CALENDAR.SLOT, calendarRange.getNumRows() - startingWeekRow + 1, p.slotsNames.length)])
    .build();
  rules.push(rule);

  // Too many slots taken compared to the reserved ones
  // TODO fix formula & divide for the different kind of slots
  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`
      =OR(
        VLOOKUP(D2; INDIRECT("${PEOPLE_SHEET_NAME}!"&A2:A); 4; FALSE) > VLOOKUP(D2; INDIRECT("${PEOPLE_SHEET_NAME}!"&A2:A); 5; FALSE);
        VLOOKUP(D2; INDIRECT("${PEOPLE_SHEET_NAME}!"&A2:A); 8; FALSE) > VLOOKUP(D2; INDIRECT("${PEOPLE_SHEET_NAME}!"&A2:A); 9; FALSE)
      )`) // Check if column D > E or H > I
    .setFontColor("red")
    .setStrikethrough(true)
    .setRanges([calendarSheet.getRange(startingWeekRow, weekCol + CALENDAR.SLOT, calendarRange.getNumRows() - startingWeekRow + 1, p.slotsNames.length)])
    .build();
  rules.push(rule);

  calendarSheet.setConditionalFormatRules(rules);
  log(`Conditional formal rules updated.`);

  // -- Make sure all pending changes are applied
  SpreadsheetApp.flush();

  info("Mise à jour terminée !");
}


/**
 * Only update people without generating the calendar again.
 */
function updatePublicPeopleOnly() {
  let publicSpreadsheet = SpreadsheetApp.openById(PUBLIC_CALENDAR_SHEET_ID);
  if (!publicSpreadsheet) {
    err(`Impossible d'ouvrir le calendrier public.`);
    return;
  }

  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let peopleActiveSheet = activeSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
  let peopleNames = getFlatDisplayValues(peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 1, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS));

  updatePublicPeople(activeSpreadsheet, publicSpreadsheet, peopleNames);

  // -- Make sure all pending changes are applied
  SpreadsheetApp.flush();

  info("Mise à jour terminée !");
}


/**
 * Update people's data on the public spreadsheet using the active's one.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} activeSpreadsheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} publicSpreadsheet
 * @param {String[]} peopleNames
 */
function updatePublicPeople(activeSpreadsheet, publicSpreadsheet, peopleNames) {
  info("Mise à jour de la liste des inscrits.");

  let parametersSheet = activeSpreadsheet.getSheetByName(PARAMETERS_SHEET_NAME);
  let freeSlotCell = parametersSheet.getRange(4, 2).getCell(1, 1);
  let unavailableSlotCell = parametersSheet.getRange(5, 2).getCell(1, 1);

  /** @type {string[][]} */
  let publicValues = Array();
  publicValues.push(["Noms"]);

  publicValues.push([freeSlotCell.getDisplayValue()]);

  publicValues.push([unavailableSlotCell.getDisplayValue()]);

  // Get people's names
  for (let name of peopleNames) {
    publicValues.push([formatName(name)]);
  }

  let peoplePublicSheet = publicSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
  peoplePublicSheet.clearContents();
  peoplePublicSheet.setRowHeights(2, peoplePublicSheet.getMaxRows() - 2, 23);

  // Add rows if there isn't enough
  if (publicValues.length > peoplePublicSheet.getMaxRows()) {
    log(`Adding ${publicValues.length - peoplePublicSheet.getMaxRows()} rows to the public people sheets.`)
    peoplePublicSheet.insertRows(peoplePublicSheet.getMaxRows(), publicValues.length - peoplePublicSheet.getMaxRows());
  }

  // TODO Add columns for Past/Future/Total/Inscrit and fill the past in the generate

  let publicRange = peoplePublicSheet.getRange(1, 1, publicValues.length)
    .setValues(publicValues)
    .setVerticalAlignment("middle");
}

/**
 * Update people's data on the public spreadsheet using the active's one.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} activeSpreadsheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} publicSpreadsheet
 * @param {String[]} peopleNames
 */
function updateActivePeople(activeSpreadsheet, publicSpreadsheet, peopleNames) {
  // TODO Recreate people list:
  //  - auto categories based on named columns categories, no more TYPES_OF_PEOPLE)
  //  - One column for each OPENING_TYPE
  //  - Past/Future/Total/Inscrit
  //  - Update Future here, not in generate
  // Line height 23
  // Vertical middle

  let peopleActiveSheet = activeSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
  let peoplePublicSheet = publicSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
}


/**
 * Add the header to `calendarSheet`.
 * @param {GenerateParameters} p
 * @param {GoogleAppsScript.Spreadsheet.Sheet} calendarSheet
 * @param {number} weekRow
 * @param {number} weekCol
 * @returns {number} The number of rows created for the header
 */
function addHeaderToCalendar(p, calendarSheet, weekRow, weekCol) {
  // -- Update the number of cols based on the nb of slots
  if (calendarSheet.getMaxColumns() > CALENDAR.SLOT + p.slotsNames.length) {
    calendarSheet.deleteColumns(
      CALENDAR.SLOT + p.slotsNames.length,
      calendarSheet.getMaxColumns() - (CALENDAR.SLOT + p.slotsNames.length)
    );
  } else if (calendarSheet.getMaxColumns() < CALENDAR.SLOT + p.slotsNames.length) {
    calendarSheet.insertColumns(
      calendarSheet.getMaxColumns(),
      CALENDAR.SLOT + p.slotsNames.length - calendarSheet.getMaxColumns()
    );
  }

  // -- Merge first cells
  let firstRange = calendarSheet.getRange(weekRow, weekCol, 1, CALENDAR.SLOT)
    .mergeAcross()
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(p.headerBackground)
    .setBorder(true, true, true, true, false, false, p.borderColor, null);
  weekCol += CALENDAR.SLOT

  // -- Set slots names
  let slotsRange = calendarSheet.getRange(weekRow, weekCol, 1, p.slotsNames.length)
    .setTextStyle(p.headerTextStyle)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(p.headerBackground)
    .setBorder(true, true, true, true, true, false, p.borderColor, null)
    .setValues([p.slotsNames])

  return 1;
}


/**
 * Create and return an array of `nbCols` length and fill it with `values`.
 * @template T
 * @param {number} nbCols
 * @param {T[] | T} value
 * @param {boolean} repeat Should the value (or the last in case of an array) be repeated to fill the row.
 * @returns {T[]}
 */
function createRowValues(nbCols, value = null, repeat = false) {
  /** @type {T[]} */
  let newRow = Array(nbCols);

  if (Array.isArray(value)) {
    if (nbCols < value.length)
      throw new RangeError(`Trying to create a row with not enought cols (${nbCols} given while expecting at least ${value.length}).`);

    for (let i = 0; i < value.length; i++) {
      newRow[i] = value[i];
    }

    if (repeat && value.length > 0) {
      let v = value[value.length - 1];
      for (let i = value.length; i < nbCols; i++) {
        newRow[i] = v;
      }
    }
  }
  else {
    if (repeat) {
      newRow.fill(value, 0, nbCols);
    } else {
      newRow[0] = value;
    }
  }

  return newRow;
}


/**
 * Create and return an initialized array for an opening row.
 * @param {number} nbCols
 * @param {OpeningTime} openingTime
 * @param {Date} begin
 * @param {Date} end
 * @param {GenerateParameters} p
 * @param {Map<[Date, Date], Array<String>>} savedMap
 * @returns {String[]}
 */
function createOpeningRow(nbCols, openingTime, begin, end, p, savedMap = null) {
  if (nbCols < CALENDAR.SLOT + p.slotsNames.length)
    throw new RangeError(`Trying to create an opening row with not enought cols (${nbCols} given while expecting at least ${CALENDAR.SLOT + p.slotsNames.length}).`);

  /** @type {String[]} */
  let newRow = Array(nbCols);

  newRow[CALENDAR.TYPE] = openingTime.type;

  newRow[CALENDAR.DAY] =
    openingTime.dayName
    + " " + begin.getDate()
    + "/" + (begin.getMonth() + 1).toString().padStart(2, 0)
    + "/" + (begin.getFullYear() - 2000).toString();

  newRow[CALENDAR.HOUR] =
    begin.getHours() + (begin.getMinutes() > 0 ? "h" + begin.getMinutes() : "h")
    + "-" + end.getHours() + (end.getMinutes() > 0 ? "h" + end.getMinutes() : "h");

  if (savedMap) {
    // Get data from the save
    /** @type {String[]} */
    let savedRow = null;
    for (const [[b, e], r] of savedMap) {
      // If the save overlaps a saved opening time
      if (b.getTime() < end.getTime() && e.getTime() > begin.getTime()) {
        savedRow = r;
        break;
      }
    }

    // Add the saved data and fill the other slots with FreeSlot
    if (savedRow) {
      for (let i = CALENDAR.SLOT; i < savedRow.length; i++) {
        newRow[i] = formatName(savedRow[i]);
      }
      for (let i = savedRow.length; i < p.slotsNames.length; i++) {
        newRow[i] = p.freeSlotCell.getDisplayValue();
      }
    }
    else {
      for (let i = 0; i < p.slotsNames.length; i++) {
        newRow[CALENDAR.SLOT + i] = p.freeSlotCell.getDisplayValue();
      }
    }
  }
  else {
    for (let i = 0; i < p.slotsNames.length; i++) {
      newRow[CALENDAR.SLOT + i] = p.freeSlotCell.getDisplayValue();
    }
  }

  return newRow;
}
