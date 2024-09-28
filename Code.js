/**
 * This App Script update the public calendar spreadsheet with the planning tables for the current week and the following ones.
 * On a new week, it removes the past one.
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

const CATEGORY_PEOPLE_COLUMNS = ["Passé", "Futur", "Total", "Payé", "Passé", "Futur", "Total", "Payé"];
const PUBLIC_CATEGORY_PEOPLE_COLUMNS = ["Total", "Payé", "Total", "Payé"];

const SELF_DAYS_HEADER = "Zone Libre";

const CALENDAR_HEADER_NB_ROWS = 2;
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
  let p = new GenerateParameters(activeSpreadsheet, true);

  let weekCol = 1;

  // -- Set wip message
  let messageRange = calendarSheet.getRange("F1");
  messageRange.setValue("MISE À JOUR EN COURS, PATIENTEZ");
  let errorRange = activeSpreadsheet.getSheetByName(HOMEPAGE_SHEET_NAME).getRange(3, 2);
  errorRange.setValue("MISE À JOUR EN COURS, ne PAS fermer la page");

  // -- Update people public list
  let peopleActiveSheet = activeSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
  let peoplePublicSheet = publicSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
  let parametersSheet = activeSpreadsheet.getSheetByName(PARAMETERS_SHEET_NAME);
  updateActivePeople(peopleActiveSheet, p.categoriesSlots);
  updatePublicPeopleNames(peopleActiveSheet, peoplePublicSheet, parametersSheet);

  // -- Update the number of cols based on the nb of slots
  let slotsNames = [];
  for (let [category, slots] of p.categoriesSlots) {
    for (let s of slots) {
      slotsNames.push(s);
    }
  }

  if (calendarSheet.getMaxColumns() > CALENDAR.SLOT + slotsNames.length) {
    calendarSheet.deleteColumns(
      CALENDAR.SLOT + slotsNames.length,
      calendarSheet.getMaxColumns() - (CALENDAR.SLOT + slotsNames.length)
    );
  } else if (calendarSheet.getMaxColumns() < CALENDAR.SLOT + slotsNames.length) {
    calendarSheet.insertColumns(
      calendarSheet.getMaxColumns(),
      CALENDAR.SLOT + slotsNames.length - calendarSheet.getMaxColumns()
    );
  }

  info("Mise à jour démarrée, ne pas fermer la page.");

  // -- Copy people data from currently displayed weeks
  let calendarRange = calendarSheet.getRange(
    CALENDAR_HEADER_NB_ROWS + 1, weekCol,
    calendarSheet.getMaxRows() - CALENDAR_HEADER_NB_ROWS, calendarSheet.getMaxColumns() - weekCol + 1
  );

  // Ensure the saveSheet has enough rows and columns
  if (saveSheet.getMaxRows() < calendarRange.getNumRows()) {
    saveSheet.insertRowsAfter(saveSheet.getMaxRows(), calendarRange.getNumRows() - saveSheet.getMaxRows());
  }
  if (saveSheet.getMaxColumns() < calendarRange.getNumColumns()) {
    saveSheet.insertColumnsAfter(saveSheet.getMaxColumns(), calendarRange.getNumColumns() - saveSheet.getMaxColumns());
  }

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

    let weekRow = CALENDAR_HEADER_NB_ROWS + 1;

    // Dropdown validation rules
    let peoplePublicSheet = publicSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
    let peopleNamePublicRange = peoplePublicSheet.getRange(2, 1, peoplePublicSheet.getMaxRows() - 1);
    let peopleRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(peopleNamePublicRange)
      .setAllowInvalid(false)
      .build();

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
    const nbCols = calendarSheet.getMaxColumns() - weekCol + 1;
    log(`Start adding weeks with ${nbCols} cols.`);
    let year = p.today.getFullYear();
    let weekNo = getWeekNumber(p.today);
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
        newDataValidations.push(createRowValues(nbCols, [null, null, null, peopleRule], true));

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
        newDataValidations.push(createRowValues(nbCols, [null, null, null, peopleRule], true));

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

    // -- Create header
    log(`Add header.`);
    let firstRange = calendarSheet.getRange(CALENDAR_HEADER_NB_ROWS, weekCol, 1, CALENDAR.SLOT)
      .mergeAcross()
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground(p.headerBackground)
      .setBorder(true, true, true, true, false, false, p.borderColor, null);

    let slotsRange = calendarSheet.getRange(CALENDAR_HEADER_NB_ROWS, weekCol + CALENDAR.SLOT, 1, slotsNames.length)
      .setTextStyle(p.headerTextStyle)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground(p.headerBackground)
      .setBorder(true, true, true, true, true, false, p.borderColor, null)
      .setValues([slotsNames])

    // -- Weeks subheader and separators
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

    // -- Update people past days counts (last so we don't count them if something went bad)
    if (UPDATE_PEOPLE_PAST_DAYS) {
      log(`Incrementing people past days.`);

      let peopleNamesActiveRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 1, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
      let peopleNamesRows = peopleNamesActiveRange.getDisplayValues();

      /** @type {Map<String, Number>} */
      let peopleToRow = new Map();
      peopleNamesRows.forEach((v, i) => peopleToRow.set(formatName(v), i));

      /** @type {Map<String, Number>} */
      let categoriesSize = new Map();

      /** @type {Map<String, GoogleAppsScript.Spreadsheet.Range>} */
      let pastDaysActiveRange = new Map();
      /** @type {Map<String, Number[][]>} */
      let pastDaysActiveValues = new Map();
      /** @type {Map<String, GoogleAppsScript.Spreadsheet.Range>} */
      let pastSelfDaysActiveRange = new Map();
      /** @type {Map<String, Number[][]>} */
      let pastSelfDaysActiveValues = new Map();

      let categoryStartCol = 2;
      for (let [category, slots] of p.categoriesSlots) {
        categoriesSize.set(category, slots.length);

        let pastDaysRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
        pastDaysActiveRange.set(category, pastDaysRange);
        pastDaysActiveValues.set(category, getValuesAsNumber(pastDaysRange));

        let pastSelfDaysRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 4, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
        pastSelfDaysActiveRange.set(category, pastSelfDaysRange);
        pastSelfDaysActiveValues.set(category, getValuesAsNumber(pastSelfDaysRange));

        categoryStartCol += 8;
      }

      for (const [[b, e], r] of savedDaysMap) {
        if (e.getTime() <= p.today.getTime()) {
          let col = CALENDAR.SLOT;

          for (let [category, slots] of p.categoriesSlots) {
            for (let i = 0; i < slots.length; i++) {
              let name = formatName(r[col + i]);
              if (name == "" || name == p.freeSlotCell.getDisplayValue() || name == p.unavailableSlotCell.getDisplayValue())
                continue;

              let row = peopleToRow.get(name);
              pastDaysActiveValues.get(category)[row][0] += 1;
            }

            col += slots.length;
          }
        }
      }
      for (const [[b, e], r] of savedSelfDaysMap) {
        if (e.getTime() <= p.today.getTime()) {
          let col = CALENDAR.SLOT;

          for (let [category, slots] of p.categoriesSlots) {
            for (let i = 0; i < slots.length; i++) {
              let name = formatName(r[col + i]);
              if (name == "" || name == p.freeSlotCell.getDisplayValue() || name == p.unavailableSlotCell.getDisplayValue())
                continue;

              let row = peopleToRow.get(name);
              pastSelfDaysActiveValues.get(category)[row][0] += 1;
            }

            col += slots.length;
          }
        }
      }

      for (let [category, slots] of p.categoriesSlots) {
        pastDaysActiveRange.get(category).setValues(pastDaysActiveValues.get(category));
        pastSelfDaysActiveRange.get(category).setValues(pastSelfDaysActiveValues.get(category));
      }
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
  let calendarSlotsRange = calendarSheet.getRange(
    CALENDAR_HEADER_NB_ROWS + 1,
    weekCol + CALENDAR.SLOT,
    calendarRange.getNumRows() - CALENDAR_HEADER_NB_ROWS,
    calendarSheet.getMaxColumns() - weekCol - CALENDAR.SLOT + 1);

  // Free slots
  let rules = [];
  let rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(p.freeSlotCell.getDisplayValue())
    .setFontColor(p.freeSlotCell.getFontColorObject().asRgbColor().asHexString())
    .setBackground(p.freeSlotCell.getBackground())
    .setRanges([calendarSlotsRange])
    .build();
  rules.push(rule);

  // Unavailable slots
  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(p.unavailableSlotCell.getDisplayValue())
    .setFontColor(p.unavailableSlotCell.getFontColorObject().asRgbColor().asHexString())
    .setBackground(p.unavailableSlotCell.getBackground())
    .setRanges([calendarSlotsRange])
    .build();
  rules.push(rule);

  // Too many slots taken compared to the reserved ones
  // TODO fix formula & divide for the different kind of slots
  // TODO Add the same formula on the name in people sheets
  // rule = SpreadsheetApp.newConditionalFormatRule()
  //   .whenFormulaSatisfied(`
  //     =OR(
  //       VLOOKUP(D2; INDIRECT("${PEOPLE_SHEET_NAME}!"&A2:A); 4; FALSE) > VLOOKUP(D2; INDIRECT("${PEOPLE_SHEET_NAME}!"&A2:A); 5; FALSE);
  //       VLOOKUP(D2; INDIRECT("${PEOPLE_SHEET_NAME}!"&A2:A); 8; FALSE) > VLOOKUP(D2; INDIRECT("${PEOPLE_SHEET_NAME}!"&A2:A); 9; FALSE)
  //     )`) // Check if column Total > Paid for all categories
  //   .setFontColor("red")
  //   .setStrikethrough(true)
  //   .setRanges([calendarSlotsRange])
  //   .build();
  // rules.push(rule);

  calendarSheet.setConditionalFormatRules(rules);
  log(`Conditional formal rules updated.`);

  // -- Update public people categories (only now because we need the past days to have been updated)
  updatePublicPeopleCategories(peopleActiveSheet, peoplePublicSheet, p.categoriesSlots);

  // -- Make sure all pending changes are applied
  SpreadsheetApp.flush();

  info("Mise à jour terminée !");
}


/**
 * Only update people without generating the calendar again.
 */
function updatePeopleOnly() {
  let publicSpreadsheet = SpreadsheetApp.openById(PUBLIC_CALENDAR_SHEET_ID);
  if (!publicSpreadsheet) {
    err(`Impossible d'ouvrir le calendrier public.`);
    return;
  }

  info("Mise à jour des tableaux d'inscrits.");

  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Categories
  let categoriesSheet = activeSpreadsheet.getSheetByName(CATEGORIES_SHEET_NAME);
  let categoriesNames = getFlatDisplayValues(categoriesSheet.getRange(2, 1, categoriesSheet.getMaxRows() - 1, 1));
  /** @type {Map<String, String[]>} */
  let categoriesSlots = new Map();
  for (let i = 0; i < categoriesNames.length; i++) {
    let slotsValues = getFlatDisplayValues(categoriesSheet.getRange(2 + i, 2, 1, categoriesSheet.getMaxColumns() - 1));
    categoriesSlots.set(categoriesNames[i], slotsValues);
  }

  let peopleActiveSheet = activeSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
  let peoplePublicSheet = publicSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
  let parametersSheet = activeSpreadsheet.getSheetByName(PARAMETERS_SHEET_NAME);

  updateActivePeople(peopleActiveSheet, categoriesSlots);
  updatePublicPeopleNames(peopleActiveSheet, peoplePublicSheet, parametersSheet);
  updatePublicPeopleCategories(peopleActiveSheet, peoplePublicSheet, categoriesSlots);

  // -- Make sure all pending changes are applied
  SpreadsheetApp.flush();

  info("Mise à jour terminée !");
}


/**
 * Update people's formatting on the active spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} peopleActiveSheet
 * @param {Map<String, String[]>} categoriesSlots
 */
function updateActivePeople(peopleActiveSheet, categoriesSlots) {
  info(`Update active people`);

  // -- Remove lines with an empty name
  let peopleActiveValues = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 1, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS).getDisplayValues();
  let rowToDelete = PEOPLE_HEADER_NB_ROWS + 1;
  for (let peopleRow of peopleActiveValues) {
    if (peopleRow[0].trim().length == 0) {
      log(`Delete empty row ${rowToDelete}`);
      peopleActiveSheet.deleteRow(rowToDelete);
    }
    else {
      rowToDelete++;
    }
  }

  // -- Remove columns of the categories that don't exist anymore
  let categoriesRow = peopleActiveSheet.getRange(1, 2, 1, peopleActiveSheet.getMaxColumns() - 1).getDisplayValues()[0];
  let colToDelete = 2;
  for (let i = 0; i < categoriesRow.length; i += 8) {
    let category = categoriesRow[i];
    if (!categoriesSlots.has(category)) {
      log(`Delete unused active category "${category}" from columns ${colToDelete} to ${colToDelete + 8}`);
      peopleActiveSheet.deleteColumns(colToDelete, 8);
    }
    else {
      colToDelete += 8;
    }
  }

  // -- Make sure we have enought columns or create the missing ones and init them
  if (peopleActiveSheet.getMaxColumns() < 1 + 8 * categoriesSlots.size) {
    let categoryStartCol = peopleActiveSheet.getMaxColumns() + 1;

    peopleActiveSheet.insertColumnsAfter(peopleActiveSheet.getMaxColumns(), 1 + 8 * categoriesSlots.size - peopleActiveSheet.getMaxColumns());

    // Past days
    peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS)
      .setValue(0);
    peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 4, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS)
      .setValue(0);

    // Total
    let pastColName = columnToLetter(categoryStartCol);
    let futureColName = columnToLetter(categoryStartCol + 1);
    let totalRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 2, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
    let total = [];
    for (let row = 0; row < totalRange.getNumRows(); row++) {
      total.push([`=${pastColName}${PEOPLE_HEADER_NB_ROWS + row + 1} + ${futureColName}${PEOPLE_HEADER_NB_ROWS + row + 1}`]);
    }
    totalRange.setValues(total);

    pastColName = columnToLetter(categoryStartCol + 4);
    futureColName = columnToLetter(categoryStartCol + 5);
    totalRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 6, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
    total = [];
    for (let row = 0; row < totalRange.getNumRows(); row++) {
      total.push([`=${pastColName}${PEOPLE_HEADER_NB_ROWS + row + 1} + ${futureColName}${PEOPLE_HEADER_NB_ROWS + row + 1}`]);
    }
    totalRange.setValues(total);

    // Paid
    peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 3, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS)
      .setValue(0);
    peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 7, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS)
      .setValue(0);
  }

  let categoryStartCol = 2;
  let calendarStartCol = CALENDAR.SLOT;
  let rules = [];

  // -- Update each category
  for (let [category, slots] of categoriesSlots) {
    // - Header
    peopleActiveSheet.getRange(1, categoryStartCol, 1, 8)
      .mergeAcross()
      .setValue(category)
      .setFontSize(13)
      .setBackground("#6aa84f");

    peopleActiveSheet.getRange(2, categoryStartCol, 1, 4)
      .mergeAcross()
      .setValue(OPENING_TYPE.REGULAR)
      .setFontSize(12)
      .setBackground("#93c47d");
    peopleActiveSheet.getRange(2, categoryStartCol + 4, 1, 4)
      .mergeAcross()
      .setValue(OPENING_TYPE.SELF)
      .setFontSize(12)
      .setBackground("#93c47d");

    peopleActiveSheet.getRange(3, categoryStartCol, 1, 8)
      .setValues([CATEGORY_PEOPLE_COLUMNS])
      .setFontSize(12)
      .setBackground("#b6d7a8");

    // - Future days formula
    let futureDaysRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 1, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
    let futureDays = [];

    let futureSelfDaysRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 5, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
    let futureSelfDays = [];

    let lastColName = columnToLetter(calendarStartCol + slots.length);
    let selectedCols = Array.from(Array(slots.length), (_, i) => i + calendarStartCol).join(";");
    let filter = `FILTER(CHOOSECOLS('${CALENDAR_SHEET_NAME}'!$A:$${lastColName}; 1; ${selectedCols}); '${CALENDAR_SHEET_NAME}'!$A:$A = "${OPENING_TYPE.REGULAR}")`;
    let filterSelf = `FILTER(CHOOSECOLS('${CALENDAR_SHEET_NAME}'!$A:$${lastColName}; 1; ${selectedCols}); '${CALENDAR_SHEET_NAME}'!$A:$A = "${OPENING_TYPE.SELF}")`;
    for (let row = 0; row < futureDaysRange.getNumRows(); row++) {
      futureDays.push([`=COUNTIF(${filter}; $A${PEOPLE_HEADER_NB_ROWS + row + 1})`]);
      futureSelfDays.push([`=COUNTIF(${filterSelf}; $A${PEOPLE_HEADER_NB_ROWS + row + 1})`]);
    }

    futureDaysRange.setValues(futureDays);
    futureSelfDaysRange.setValues(futureSelfDays);

    // - Style

    // Past days
    peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS)
      .setBackground("#efefef")
      .setFontColor("#999999");
    peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 4, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS)
      .setBackground("#efefef")
      .setFontColor("#999999");

    // Future days
    peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 1, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS)
      .setBackground("#efefef")
      .setFontColor("#999999");
    peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 5, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS)
      .setBackground("#efefef")
      .setFontColor("#999999");

    // Total
    let totalRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 2, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS)
      .setBackground("#f3f3f3");
    let totalSelfRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryStartCol + 6, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS)
      .setBackground("#f3f3f3");

    let totalLetter = columnToLetter(categoryStartCol + 2);
    let paidLetter = columnToLetter(categoryStartCol + 3);
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(${paidLetter}${PEOPLE_HEADER_NB_ROWS + 1}<>""; ${totalLetter}${PEOPLE_HEADER_NB_ROWS + 1} > ${paidLetter}${PEOPLE_HEADER_NB_ROWS + 1})`)
      .setFontColor("#ffffff")
      .setBackground("#cc0000")
      .setRanges([totalRange])
      .build();
    rules.push(rule);

    let totalSelfLetter = columnToLetter(categoryStartCol + 6);
    let paidSelfLetter = columnToLetter(categoryStartCol + 7);
    rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(${paidSelfLetter}${PEOPLE_HEADER_NB_ROWS + 1}<>""; ${totalSelfLetter}${PEOPLE_HEADER_NB_ROWS + 1} > ${paidSelfLetter}${PEOPLE_HEADER_NB_ROWS + 1})`)
      .setFontColor("#ffffff")
      .setBackground("#cc0000")
      .setRanges([totalSelfRange])
      .build();
    rules.push(rule);

    // Borders
    peopleActiveSheet.getRange(1, categoryStartCol, peopleActiveSheet.getMaxRows(), 8)
      .setBorder(null, null, null, true, null, null, "#333333", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    peopleActiveSheet.getRange(1, categoryStartCol, peopleActiveSheet.getMaxRows(), 4)
      .setBorder(null, null, null, true, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);

    calendarStartCol += slots.length;
    categoryStartCol += 8;
  }

  peopleActiveSheet.setConditionalFormatRules(rules);

  // -- Style
  peopleActiveSheet.getRange(1, 1, PEOPLE_HEADER_NB_ROWS, peopleActiveSheet.getMaxColumns())
    .setHorizontalAlignment("center")
    .setFontWeight("bold");
  peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 1, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS, peopleActiveSheet.getMaxColumns())
    .setVerticalAlignment("middle");
  peopleActiveSheet.setColumnWidths(2, peopleActiveSheet.getMaxColumns() - 1, 60);
  peopleActiveSheet.setRowHeights(1, peopleActiveSheet.getMaxRows(), 23);
}


/**
 * Update people's data on the public spreadsheet using the active's one.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} peopleActiveSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} peoplePublicSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} parametersSheet
 */
function updatePublicPeopleNames(peopleActiveSheet, peoplePublicSheet, parametersSheet) {
  info("Update public people name");

  // Add rows if there isn't enough
  let maxRows = peopleActiveSheet.getMaxRows() + 2;
  if (maxRows > peoplePublicSheet.getMaxRows()) {
    log(`Adding ${maxRows - peoplePublicSheet.getMaxRows()} rows to the public people sheets.`)
    peoplePublicSheet.insertRowsAfter(peoplePublicSheet.getMaxRows(), maxRows - peoplePublicSheet.getMaxRows());
  } else if (maxRows < peoplePublicSheet.getMaxRows()) {
    log(`Removing ${peoplePublicSheet.getMaxRows() - maxRows} rows from the public people sheets.`)
    peoplePublicSheet.deleteRows(maxRows, peoplePublicSheet.getMaxRows() - maxRows);
  }

  // -- Names
  let peopleActiveValues = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 1, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS).getDisplayValues();
  let freeSlotCell = parametersSheet.getRange(4, 2).getCell(1, 1);
  let unavailableSlotCell = parametersSheet.getRange(5, 2).getCell(1, 1);
  let peopleValues = Array([freeSlotCell.getDisplayValue()], [unavailableSlotCell.getDisplayValue()]);
  for (let row of peopleActiveValues) {
    peopleValues.push([formatName(row[0])]);
  }

  peoplePublicSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 1, peopleValues.length)
    .setValues(peopleValues);
}


/**
 * Update people's data on the public spreadsheet using the active's one.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} peopleActiveSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} peoplePublicSheet
 * @param {Map<String, String[]>} categoriesSlots
 */
function updatePublicPeopleCategories(peopleActiveSheet, peoplePublicSheet, categoriesSlots) {
  info("Update public people list");

  // -- Remove columns of the categories that don't exist anymore
  let categoriesRow = peoplePublicSheet.getRange(1, 2, 1, peoplePublicSheet.getMaxColumns() - 1).getDisplayValues()[0];
  let colToDelete = 2;
  for (let i = 0; i < categoriesRow.length; i += 4) {
    let category = categoriesRow[i];
    if (!categoriesSlots.has(category)) {
      log(`Delete unused public category "${category}" from columns ${colToDelete} to ${colToDelete + 4}`);
      peoplePublicSheet.deleteColumns(colToDelete, 4);
    }
    else {
      colToDelete += 4;
    }
  }

  // Make sure we have enought columns or create the missing ones and init them
  if (peoplePublicSheet.getMaxColumns() < 1 + 4 * categoriesSlots.size) {
    peoplePublicSheet.insertColumnsAfter(peoplePublicSheet.getMaxColumns(), 1 + 4 * categoriesSlots.size - peoplePublicSheet.getMaxColumns());
  }

  // -- Categories
  let categoryStartCol = 2;
  let categoryActiveStartCol = 2;
  let calendarStartCol = CALENDAR.SLOT;
  let rules = [];

  for (let [category, slots] of categoriesSlots) {
    // - Header
    peoplePublicSheet.getRange(1, categoryStartCol, 1, 4)
      .mergeAcross()
      .setValue(category)
      .setFontSize(13)
      .setBackground("#6aa84f");

    peoplePublicSheet.getRange(2, categoryStartCol, 1, 2)
      .mergeAcross()
      .setValue(OPENING_TYPE.REGULAR)
      .setFontSize(12)
      .setBackground("#93c47d");
    peoplePublicSheet.getRange(2, categoryStartCol + 2, 1, 2)
      .mergeAcross()
      .setValue(OPENING_TYPE.SELF)
      .setFontSize(12)
      .setBackground("#93c47d");

    peoplePublicSheet.getRange(3, categoryStartCol, 1, 4)
      .setValues([PUBLIC_CATEGORY_PEOPLE_COLUMNS])
      .setFontSize(12)
      .setBackground("#b6d7a8");

    // - Future days formula (skip free and unavailable)
    let pastActiveValues = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryActiveStartCol, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS).getDisplayValues();
    let futureDaysRange = peoplePublicSheet.getRange(PEOPLE_HEADER_NB_ROWS + 3, categoryStartCol, peoplePublicSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS - 2);
    let futureDays = [];

    let pastSelfActiveValues = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryActiveStartCol + 4, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS).getDisplayValues();
    let futureSelfDaysRange = peoplePublicSheet.getRange(PEOPLE_HEADER_NB_ROWS + 3, categoryStartCol + 2, peoplePublicSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS - 2);
    let futureSelfDays = [];

    let lastColName = columnToLetter(calendarStartCol + slots.length);
    let selectedCols = Array.from(Array(slots.length), (_, i) => i + calendarStartCol).join(";");
    let filter = `FILTER(CHOOSECOLS('${CALENDAR_SHEET_NAME}'!$A:$${lastColName}; 1; ${selectedCols}); '${CALENDAR_SHEET_NAME}'!$A:$A = "${OPENING_TYPE.REGULAR}")`;
    let filterSelf = `FILTER(CHOOSECOLS('${CALENDAR_SHEET_NAME}'!$A:$${lastColName}; 1; ${selectedCols}); '${CALENDAR_SHEET_NAME}'!$A:$A = "${OPENING_TYPE.SELF}")`;
    for (let row = 0; row < futureDaysRange.getNumRows(); row++) {
      futureDays.push([`=${pastActiveValues[row]} + COUNTIF(${filter}; $A${PEOPLE_HEADER_NB_ROWS + 3 + row})`]);
      futureSelfDays.push([`=${pastSelfActiveValues[row]} + COUNTIF(${filterSelf}; $A${PEOPLE_HEADER_NB_ROWS + 3 + row})`]);
    }

    futureDaysRange.setValues(futureDays);
    futureSelfDaysRange.setValues(futureSelfDays);

    // - Copy paid
    let paidActiveValues = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryActiveStartCol + 3, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS).getDisplayValues();
    peoplePublicSheet.getRange(PEOPLE_HEADER_NB_ROWS + 3, categoryStartCol + 1, peoplePublicSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS - 2)
      .setValues(paidActiveValues);

    let selfpaidActiveValues = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, categoryActiveStartCol + 7, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS).getDisplayValues();
    peoplePublicSheet.getRange(PEOPLE_HEADER_NB_ROWS + 3, categoryStartCol + 3, peoplePublicSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS - 2)
      .setValues(selfpaidActiveValues);

    // - Style

    // Total
    let totalRange = peoplePublicSheet.getRange(PEOPLE_HEADER_NB_ROWS + 3, categoryStartCol, peoplePublicSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS - 2)
      .setBackground("#f3f3f3");
    let totalSelfRange = peoplePublicSheet.getRange(PEOPLE_HEADER_NB_ROWS + 3, categoryStartCol + 2, peoplePublicSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS - 2)
      .setBackground("#f3f3f3");

    let totalLetter = columnToLetter(categoryStartCol);
    let paidLetter = columnToLetter(categoryStartCol + 1);
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(${paidLetter}${PEOPLE_HEADER_NB_ROWS + 3}<>""; ${totalLetter}${PEOPLE_HEADER_NB_ROWS + 3} > ${paidLetter}${PEOPLE_HEADER_NB_ROWS + 3})`)
      .setFontColor("#ffffff")
      .setBackground("#cc0000")
      .setRanges([totalRange])
      .build();
    rules.push(rule);

    let totalSelfLetter = columnToLetter(categoryStartCol + 2);
    let paidSelfLetter = columnToLetter(categoryStartCol + 3);
    rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(${paidSelfLetter}${PEOPLE_HEADER_NB_ROWS + 3}<>""; ${totalSelfLetter}${PEOPLE_HEADER_NB_ROWS + 3} > ${paidSelfLetter}${PEOPLE_HEADER_NB_ROWS + 3})`)
      .setFontColor("#ffffff")
      .setBackground("#cc0000")
      .setRanges([totalSelfRange])
      .build();
    rules.push(rule);

    // Borders
    peoplePublicSheet.getRange(1, categoryStartCol, peoplePublicSheet.getMaxRows(), 4)
      .setBorder(null, null, null, true, null, null, "#333333", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    peoplePublicSheet.getRange(1, categoryStartCol, peoplePublicSheet.getMaxRows(), 2)
      .setBorder(null, null, null, true, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);

    calendarStartCol += slots.length;
    categoryStartCol += 4;
    categoryActiveStartCol += 8;
  }

  peoplePublicSheet.setConditionalFormatRules(rules);

  // - Style
  peoplePublicSheet.getRange(1, 1, PEOPLE_HEADER_NB_ROWS, peoplePublicSheet.getMaxColumns())
    .setHorizontalAlignment("center")
    .setFontWeight("bold");
  peoplePublicSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 1, peoplePublicSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS, peoplePublicSheet.getMaxColumns())
    .setVerticalAlignment("middle");
  peoplePublicSheet.setColumnWidths(2, peoplePublicSheet.getMaxColumns() - 1, 60);
  peoplePublicSheet.setRowHeights(1, peoplePublicSheet.getMaxRows(), 23);
}


function addOnePerson() {
  let publicSpreadsheet = SpreadsheetApp.openById(PUBLIC_CALENDAR_SHEET_ID);
  if (!publicSpreadsheet) {
    err(`Impossible d'ouvrir le calendrier public.`);
    return;
  }
  log(`Adding one person`);

  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    "Nom du nouvel inscrit:",
    ui.ButtonSet.OK_CANCEL,
  );

  var button = result.getSelectedButton();
  if (button == ui.Button.OK) {
  } else if (button == ui.Button.CANCEL) {
    log(`Canceled`);
    return;
  } else if (button == ui.Button.CLOSE) {
    log(`Closed`);
    return;
  }

  var name = result.getResponseText();

  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let peopleActiveSheet = activeSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
  let peoplePublicSheet = publicSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
  let parametersSheet = activeSpreadsheet.getSheetByName(PARAMETERS_SHEET_NAME);

  // Add the line and fill the name
  peopleActiveSheet.insertRowAfter(peopleActiveSheet.getMaxRows());
  let newPersonRange = peopleActiveSheet.getRange(peopleActiveSheet.getMaxRows(), 1, 1, peopleActiveSheet.getMaxColumns());
  let newPersonRow = [name];
  for (let i = 1; i < peopleActiveSheet.getMaxColumns(); i += 4) {
    // Past
    newPersonRow.push(0);

    // Future (initialized in `updateActivePeople`)
    newPersonRow.push(0);

    // Total
    let pastColName = columnToLetter(i + 1);
    let futureColName = columnToLetter(i + 2);
    newPersonRow.push(`=${pastColName}${peopleActiveSheet.getMaxRows()} + ${futureColName}${peopleActiveSheet.getMaxRows()}`);

    // Paid
    newPersonRow.push(0);
  }
  newPersonRange.setValues([newPersonRow]);

  // Categories
  let categoriesSheet = activeSpreadsheet.getSheetByName(CATEGORIES_SHEET_NAME);
  let categoriesNames = getFlatDisplayValues(categoriesSheet.getRange(2, 1, categoriesSheet.getMaxRows() - 1, 1));
  /** @type {Map<String, String[]>} */
  let categoriesSlots = new Map();
  for (let i = 0; i < categoriesNames.length; i++) {
    let slotsValues = getFlatDisplayValues(categoriesSheet.getRange(2 + i, 2, 1, categoriesSheet.getMaxColumns() - 1));
    categoriesSlots.set(categoriesNames[i], slotsValues);
  }

  updateActivePeople(peopleActiveSheet, categoriesSlots);
  updatePublicPeopleNames(peopleActiveSheet, peoplePublicSheet, parametersSheet);
  updatePublicPeopleCategories(peopleActiveSheet, peoplePublicSheet, categoriesSlots);

  // -- Make sure all pending changes are applied
  SpreadsheetApp.flush();

  info(`${name} a bien été ajouté !`);
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
 * @param {Date} begin The date corresponding to the beginning of this specific opening.
 * @param {Date} end The date corresponding to the end of this specific opening.
 * @param {GenerateParameters} p
 * @param {Map<[Date, Date], Array<String>>} savedMap
 * @returns {String[]}
 */
function createOpeningRow(nbCols, openingTime, begin, end, p, savedMap = null) {
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
      for (let i = CALENDAR.SLOT; i < Math.min(nbCols, savedRow.length); i++) {
        newRow[i] = formatName(savedRow[i]);
        if (newRow[i] == "")
          newRow[i] = p.freeSlotCell.getDisplayValue();
      }
      for (let i = savedRow.length; i < nbCols; i++) {
        newRow[i] = p.freeSlotCell.getDisplayValue();
      }
    }
    else {
      for (let i = CALENDAR.SLOT; i < nbCols; i++) {
        newRow[i] = p.freeSlotCell.getDisplayValue();
      }
    }
  }
  else {
    for (let i = CALENDAR.SLOT; i < nbCols; i++) {
      newRow[i] = p.freeSlotCell.getDisplayValue();
    }
  }

  return newRow;
}
