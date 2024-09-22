/**
 * This App Script update the calendar sheet with the planning tables for the current week and the following ones.
 * On a new week, it removes the past one and adds a empty new one at the end.
 * Named ranges are used to parameter all.
 */

const APP_TITLE = "Agenda Atelier Nuances";

const CALENDAR_SHEET_NAME = "Calendrier Céramistes";
const PEOPLE_SHEET_NAME = "Inscrits";
const SAVE_SHEET_NAME = "SaveData";

const PUBLIC_CALENDAR_SHEET_ID = "1_0Mh-E4UW4-eC-Y6oMr3VwMoKZSNE-i3woWh5OpXDmA";

/** Should the data in the calendar be conserved. If false, it will be generated empty. */
const KEEP_CALENDAR_DATA = true;
/** Should people past days counts be updated by adding the days removed and prior to today. */
const UPDATE_PEOPLE_PAST_DAYS = true;

const TYPES_OF_PEOPLE = ['Tourneurs', 'Modeleurs'];

const SELF_DAYS_HEADER = "Zone Libre";


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
  let weekRow = startingWeekRow + 1;
  let weekCol = 1;

  // -- Set wip message
  let messageRange = calendarSheet.getRange("F1");
  messageRange.setValue("MISE À JOUR EN COURS, PATIENTEZ");
  let errorRange = activeSpreadsheet.getRangeByName("Erreurs");
  errorRange.setValue("MISE À JOUR EN COURS, ne PAS fermer la page");

  info("Mise à jour démarrée, ne pas fermer la page.");

  // -- Update people list
  updatePeople(activeSpreadsheet, publicSpreadsheet, p.peopleNames);

  // -- Copy people data from currently displayed weeks
  let calendarRange = calendarSheet.getRange(
    startingWeekRow, weekCol,
    calendarSheet.getMaxRows() - startingWeekRow + 1, calendarSheet.getMaxColumns() - weekCol + 1
  );
  calendarRange.clearDataValidations();
  calendarRange.setNumberFormat("@");

  let savedValues = calendarRange.getDisplayValues();
  let saveRange = saveSheet.getRange(1, 1, calendarRange.getNumRows(), calendarRange.getNumColumns());
  saveRange.clear();
  saveRange.clearDataValidations();
  saveRange.setNumberFormat("@");

  // Copy to save sheet
  saveRange.setValues(savedValues);

  // Generate saved values map
  log("Generate saved values map");
  /** @type {Map<[Date, Date], Array<String>>} */
  let savedDaysMap = Map();
  /** @type {Map<[Date, Date], Array<String>>} */
  let savedSelfDaysMap = Map();
  let isSelfDay = false;
  for (let row = 0; row < savedValues.length; row++) {

    // Update isSelfDay and skip empty rows
    if (newValues[row][CALENDAR.HOUR] == "") {
      isSelfDay = (newValues[row][0] == SELF_DAYS_HEADER);
      continue;
    }

    let savedRow = savedValues[row];
    let savedDate = new Date();

    // Get begin/end date
    let savedDaySplit = savedRow[CALENDAR.DAY].trim().split(" ");
    if (savedDaySplit.length > 1) {
      updateDate(savedDate, savedDaySplit[savedDaySplit.length - 1]);
    } else {
      continue;
    }

    let savedEndDate = new Date();
    let savedHourSplit = savedRow[CALENDAR.HOUR].split("-");
    if (savedHourSplit.length > 1) {
      updateTime(savedDate, savedHourSplit[0]);

      savedEndDate.setFullYear(savedDate.getFullYear(), savedDate.getMonth(), savedDate.getDate());
      updateTime(savedEndDate, savedHourSplit[1]);
    } else {
      continue;
    }

    if (isSelfDay) {
      savedSelfDaysMap.set([savedDate, savedEndDate], savedRow);
    } else {
      savedDaysMap.set([savedDate, savedEndDate], savedRow);
    }
  }

  info("Calendrier sauvegardé");

  // -- Clear
  calendarRange.clear();
  log(`Cleared the public calendar.`);

  calendarSheet.setRowHeights(1, calendarSheet.getMaxRows(), 21);

  // -- Try to update the sheet and if there is an issue, copy back the saved sheet
  try {
    let weekNo = getWeekNumber(p.today);

    // -- Create header
    addHeaderToCalendar(p, calendarSheet, startingWeekRow, weekCol);
    log(`Added header.`);

    // -- Add weeks calendars, starting from current
    let year = p.today.getFullYear();
    for (let weekIdx = 0; weekIdx < p.weeksToDisplay; weekIdx++) {
      if (weekNo + weekIdx > 52) {
        year++;
        weekNo -= 52;
      }
      weekRow += addWeekToCalendar(p, weekNo + weekIdx, year, calendarSheet, weekRow, weekCol);
      log(`Added week ${weekNo + weekIdx} of ${year}.`)
    }

    // -- Insert back saved people data
    if (KEEP_CALENDAR_DATA) {
      let newValues = calendarRange.getDisplayValues();
      let saveRow = 1;
      let savedDate = new Date();
      let savedEndDate = new Date();
      let newDate = new Date();
      let newEndDate = new Date();

      for (let row = 0; row < newValues.length; row++) {
        // We browse two arrays at the same time:
        // `savedValues` which use `saveRow` as an index
        // `newValues` which use `row` as an index

        // We skip headers which only have their first cell filled
        if (newValues[row][CALENDAR.HOUR] == "")
          continue;

        // We skip saved rows with empty slots
        while (saveRow < savedValues.length && savedValues[saveRow][CALENDAR.HOUR] == "") {
          saveRow++;
        }

        // All saved data have been inserted
        if (saveRow >= savedValues.length)
          break;

        // If the new row does not match the saved one, there was a deletion or addition
        if (newValues[row][CALENDAR.DAY] != savedValues[saveRow][CALENDAR.DAY]
          || newValues[row][CALENDAR.HOUR] != savedValues[saveRow][CALENDAR.HOUR]) {
          // Check the saved row:
          // If the saved date is further in the future than the new one, a new row was added
          // and we just don't increase `saveRow` as we wait for `row` to catch up.
          // If the saved date is further in the past than the new one, the saved row was deleted
          // and we must look for the next matching saved row.

          let savedRow = savedValues[saveRow];

          let savedDaySplit = savedRow[CALENDAR.DAY].trim().split(" ");
          if (savedDaySplit.length > 1)
            updateDate(savedDate, savedDaySplit[savedDaySplit.length - 1]);
          let savedHourSplit = savedRow[CALENDAR.HOUR].split("-");
          if (savedHourSplit.length > 1)
            updateTime(savedDate, savedHourSplit[0]);

          let newRow = newValues[row];

          let newDaySplit = newRow[CALENDAR.DAY].trim().split(" ");
          if (newDaySplit.length > 1)
            updateDate(newEndDate, newDaySplit[newDaySplit.length - 1]);
          let newHourSplit = newRow[CALENDAR.HOUR].split("-");
          if (newHourSplit.length > 1)
            updateTime(newEndDate, newHourSplit[1]);

          // If the saved date is further in the future than the new one, a new row was added.
          // We continue and don't increase `saveRow` as we wait for `row` to catch up.
          if (savedDate >= newEndDate)
            continue;

          if (savedHourSplit.length > 1) {
            savedEndDate.setFullYear(savedDate.getFullYear(), savedDate.getMonth(), savedDate.getDate());
            updateTime(savedEndDate, savedHourSplit[1]);
          }

          if (newHourSplit.length > 1) {
            newDate.setFullYear(newEndDate.getFullYear(), newEndDate.getMonth(), newEndDate.getDate());
            updateTime(newDate, newHourSplit[0]);
          }

          // If the saved date is further in the past than the new one, the saved row was deleted.
          // We must look for the next matching saved row and discard the old ones.
          if (savedEndDate <= newDate) {
            let saveIsOlder = false;

            for (let r = saveRow + 1; r < savedValues.length; r++) {

              // Increase the people past days counter if the saved date is older than today
              if (savedDate < p.today) {
                for (let s = 0; s < p.ceramistsSlotsName.length; s++) {
                  p.addPastDay(savedValues[r - 1][CALENDAR.SLOT + s], s);
                }

                for (let s = 0; s < p.modelersSlotsName.length; s++) {
                  p.addPastDay(savedValues[r - 1][CALENDAR.SLOT + s], s);
                }
                log(`Past day ${savedDate}.`)
              }

              // Skip headers and rows of the same day/hour that list the slots
              if (savedValues[r][CALENDAR.HOUR] == "")
                continue;

              savedRow = savedValues[r];

              savedDaySplit = savedRow[CALENDAR.DAY].trim().split(" ");
              if (savedDaySplit.length > 1)
                updateDate(savedDate, savedDaySplit[savedDaySplit.length - 1]);
              savedHourSplit = savedRow[CALENDAR.HOUR].trim().split("-");
              if (savedHourSplit.length > 1) {
                updateTime(savedDate, savedHourSplit[0]);

                savedEndDate.setFullYear(savedDate.getFullYear(), savedDate.getMonth(), savedDate.getDate());
                updateTime(savedEndDate, savedHourSplit[1]);
              }

              // Does the save date now overlap with the new one?
              if (savedDate < newEndDate && savedEndDate > newDate) {
                saveRow = r;
                break;
              }

              // Is the save date now further in the future than the new date?
              if (savedDate >= newEndDate) {
                saveRow = r;
                saveIsOlder = true;
                break;
              }
            }
            // The next save date is further in the future than the new one.
            // We continue and don't increase `saveRow` anymore as we wait for `row` to catch up.
            if (saveIsOlder)
              continue;

            // Now the saved date overlap with the new one so we can take data from the save.
          }
        }

        // Copy from the save
        for (let slot = 0; slot < p.slotsNames.length; slot++) {
          if (savedValues[saveRow][CALENDAR.SLOT + slot] != "")
            newValues[row][CALENDAR.SLOT + slot] = formatName(savedValues[saveRow][CALENDAR.SLOT + slot]);
        }
        saveRow++;
      }

      calendarRange.setValues(newValues);
      log(`Calendar data restored.`);
    }

    // -- Update people past days counts
    if (UPDATE_PEOPLE_PAST_DAYS) {
      p.ceramistsPastDaysActiveRange.setValues(p.ceramistsPastDays);
      p.ceramistsSelfPastDaysActiveRange.setValues(p.ceramistsSelfPastDays);
      p.modelersPastDaysActiveRange.setValues(p.modelersPastDays);
      p.modelersSelfPastDaysActiveRange.setValues(p.modelersSelfPastDays);
      log(`People past days updated.`);

      // TODO Update people "days to come" formula in the active people sheet.
      // Formula is =COUNTIF('Calendrier Céramistes'!$D:$E; $A4) but $D:$E needs
      // to be changed depending on the category (ceramists/modelers) and $A4 should match the curent row.
    }

    errorRange.clearContent();
  } catch (e) {
    err(`Erreur pendant l'insertion des valeurs, sauvegarde restaurée.`, e);
    calendarRange.setValues(savedValues);
    // saveRange.copyTo(calendarRange);
    errorRange.setValue(`Erreur prevenir Grégoire, ne rien toucher.`);
  }

  messageRange.clearContent();

  // -- Set conditional format rules
  // Free slots
  let rules = [];
  let rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(p.freeSlotCell.getDisplayValue())
    .setFontColor(p.freeSlotCell.getFontColorObject().asRgbColor().asHexString())
    .setBackground(p.freeSlotCell.getBackground())
    .setRanges([calendarSheet.getRange(startingWeekRow, weekCol + CALENDAR.SLOT, calendarRange.getNumRows(), p.slotsNames.length)])
    .build();
  rules.push(rule);

  // Unavailable slots
  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(p.unavailableSlotCell.getDisplayValue())
    .setFontColor(p.unavailableSlotCell.getFontColorObject().asRgbColor().asHexString())
    .setBackground(p.unavailableSlotCell.getBackground())
    .setRanges([calendarSheet.getRange(startingWeekRow, weekCol + CALENDAR.SLOT, calendarRange.getNumRows(), p.slotsNames.length)])
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
    .setRanges([calendarSheet.getRange(startingWeekRow, weekCol + CALENDAR.SLOT, calendarRange.getNumRows(), p.slotsNames.length)])
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
function updatePeopleOnly() {
  let publicSpreadsheet = SpreadsheetApp.openById(PUBLIC_CALENDAR_SHEET_ID);
  if (!publicSpreadsheet) {
    err(`Impossible d'ouvrir le calendrier public.`);
    return;
  }

  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let peopleActiveSheet = activeSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
  let peopleNames = getFlatDisplayValues(peopleActiveSheet.getRange(4, 1, peopleActiveSheet.getMaxRows()));

  updatePeople(activeSpreadsheet, publicSpreadsheet, peopleNames);

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
function updatePeople(activeSpreadsheet, publicSpreadsheet, peopleNames) {
  info("Mise à jour de la liste des inscrits.");

  let freeSlotCell = activeSpreadsheet.getRangeByName('EmplacementLibre').getCell(1, 1);
  let unavailableSlotCell = activeSpreadsheet.getRangeByName('EmplacementIndisponible').getCell(1, 1);

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

  // Add rows if there isn't enough
  if (publicValues.length > peoplePublicSheet.getMaxRows()) {
    log(`Adding ${publicValues.length - peoplePublicSheet.getMaxRows()} rows to the public people sheets.`)
    peoplePublicSheet.insertRows(peoplePublicSheet.getMaxRows(), publicValues.length - peoplePublicSheet.getMaxRows());
  }

  let publicRange = peoplePublicSheet.getRange(1, 1, publicValues.length);
  publicRange.setValues(publicValues);
}

/**
 * Add the formated week to calendarSheet and returns the number of edited rows.
 * @param {GenerateParameters} p
 * @param {GoogleAppsScript.Spreadsheet.Sheet} calendarSheet
 * @param {number} weekRow
 * @param {number} weekCol
 * @OnlyCurrentDoc */
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
  let firstRange = calendarSheet.getRange(weekRow, weekCol, 1, CALENDAR.SLOT);
  firstRange.mergeAcross();
  firstRange.setHorizontalAlignment("center");
  firstRange.setVerticalAlignment("middle");
  firstRange.setBackground(p.headerBackground);
  firstRange.setBorder(true, true, true, true, false, false, p.borderColor, null);
  weekCol += CALENDAR.SLOT

  // -- Set slots names
  let slotsRange = calendarSheet.getRange(weekRow, weekCol, 1, p.slotsNames.length);
  slotsRange.setTextStyle(p.headerTextStyle);
  slotsRange.setHorizontalAlignment("center");
  slotsRange.setVerticalAlignment("middle");
  slotsRange.setBackground(p.headerBackground);
  slotsRange.setBorder(true, true, true, true, true, false, p.borderColor, null);
  slotsRange.setValues([p.slotsNames])
}


/**
 * Add the formated week to calendarSheet and returns the number of edited rows.
 * @param {GenerateParameters} p
 * @param {number} weekNo
 * @param {number} year
 * @param {GoogleAppsScript.Spreadsheet.Sheet} calendarSheet
 * @param {number} weekRow
 * @param {number} weekCol
 * @return {number} The number of edited rows (1 for the title + number of opening [+ subheader + number of self opening])
 * @OnlyCurrentDoc */
function addWeekToCalendar(p, weekNo, year, calendarSheet, weekRow, weekCol) {
  let nbCols = calendarSheet.getMaxColumns() - weekCol + 1;

  let dateOfWeek = getDateOfWeek(weekNo, year);

  // -- Set the title
  let titleRange = calendarSheet.getRange(weekRow, weekCol, 1, nbCols);
  titleRange.mergeAcross();
  titleRange.setValue("Semaine " + weekNo + " - " + year);
  titleRange.setTextStyle(p.headerTextStyle);
  titleRange.setHorizontalAlignment("center");
  titleRange.setVerticalAlignment("middle");
  titleRange.setBackground(p.subheaderBackground);
  titleRange.setBorder(true, true, true, true, false, false, p.borderColor, null);

  weekRow += 1;
  let currentRow = weekRow;

  // -- Generate one line per opening time

  let daysBgs = Array();
  let hoursBgs = Array();

  // Regular openings
  let openingsValues = Array();
  for (let openingTime of p.openingTimes) {
    let begin = new Date(
      dateOfWeek.getFullYear(),
      dateOfWeek.getMonth(),
      dateOfWeek.getDate() + openingTime.dayOfWeek,
      openingTime.begin.getHours(),
      openingTime.begin.getMinutes()
    );

    let end = new Date(
      dateOfWeek.getFullYear(),
      dateOfWeek.getMonth(),
      dateOfWeek.getDate() + openingTime.dayOfWeek,
      openingTime.end.getHours(),
      openingTime.end.getMinutes()
    );

    // Don't add the slots if closed
    let skipSlot = false;
    for (let closedTime of p.closedTimes) {
      if (end.getTime() >= closedTime.begin.getTime()
        && begin.getTime() <= closedTime.end.getTime()) {
        skipSlot = true;
        break;
      }
    }
    if (skipSlot)
      continue;

    // Set opening values
    let opening = Array(nbCols);
    let dayName = openingTime.dayName;
    opening[CALENDAR.TYPE] = OPENING_TYPE.REGULAR;
    opening[CALENDAR.DAY] =
      dayName.slice(0, 1).toUpperCase() + dayName.slice(1, 3)
      + " " + begin.getDate()
      + "/" + (begin.getMonth() + 1).toString().padStart(2, 0);
    opening[CALENDAR.HOUR] =
      begin.getHours() + (begin.getMinutes() > 0 ? "h" + begin.getMinutes() : "h")
      + "-" + end.getHours() + (end.getMinutes() > 0 ? "h" + end.getMinutes() : "h");
    for (let i = 0; i < p.slotsNames.length; i++) {
      opening[CALENDAR.SLOT + i] = p.freeSlotCell.getDisplayValue();
    }

    openingsValues.push(opening);

    daysBgs.push([openingTime.dayColor]);
    hoursBgs.push([openingTime.hourColor]);

    currentRow += 1;
  }

  // Skip the subheader
  daysBgs.push([null]);
  hoursBgs.push([null]);
  currentRow += 1;

  // Self-openings
  let selfopeningsValues = Array();
  for (let openingTime of p.selfopeningTimes) {
    let begin = new Date(
      dateOfWeek.getFullYear(),
      dateOfWeek.getMonth(),
      dateOfWeek.getDate() + openingTime.dayOfWeek,
      openingTime.begin.getHours(),
      openingTime.begin.getMinutes()
    );

    let end = new Date(
      dateOfWeek.getFullYear(),
      dateOfWeek.getMonth(),
      dateOfWeek.getDate() + openingTime.dayOfWeek,
      openingTime.end.getHours(),
      openingTime.end.getMinutes()
    );

    // Don't add the slots if closed
    let skipSlot = false;
    for (let closedTime of p.closedTimes) {
      if (end.getTime() >= closedTime.begin.getTime()
        && begin.getTime() <= closedTime.end.getTime()) {
        skipSlot = true;
        break;
      }
    }
    if (skipSlot)
      continue;

    // Set opening values
    let opening = Array(nbCols);
    let dayName = openingTime.dayName;
    opening[CALENDAR.TYPE] = OPENING_TYPE.SELF;
    opening[CALENDAR.DAY] =
      dayName.slice(0, 1).toUpperCase() + dayName.slice(1, 3)
      + " " + begin.getDate()
      + "/" + (begin.getMonth() + 1).toString().padStart(2, 0);
    opening[CALENDAR.HOUR] =
      begin.getHours() + (begin.getMinutes() > 0 ? "h" + begin.getMinutes() : "h")
      + "-" + end.getHours() + (end.getMinutes() > 0 ? "h" + end.getMinutes() : "h");
    for (let i = 0; i < p.slotsNames.length; i++) {
      opening[CALENDAR.SLOT + i] = p.freeSlotCell.getDisplayValue();
    }

    selfopeningsValues.push(opening);

    daysBgs.push([openingTime.dayColor]);
    hoursBgs.push([openingTime.hourColor]);

    currentRow += 1;
  }

  // -- If there is no slot available this week, skip it
  if (openingsValues.length == 0 && selfopeningsValues.length == 0)
    return 1;

  // -- Insert the values in the range of the week's calendar
  // Openings
  let openingsRange = calendarSheet.getRange(weekRow, weekCol, openingsValues.length, nbCols);
  openingsRange.setValues(openingsValues);

  // Separator
  let separatorRange = calendarSheet.getRange(weekRow + openingsValues.length, weekCol, 1, nbCols);
  separatorRange.mergeAcross();
  separatorRange.setValue(SELF_DAYS_HEADER);
  separatorRange.setTextStyle(p.subheaderTextStyle);
  separatorRange.setHorizontalAlignment("center");
  separatorRange.setVerticalAlignment("middle");
  separatorRange.setBackground(p.subheaderBackground);
  separatorRange.setBorder(true, true, true, true, false, false, p.borderColor, null);

  // Self-openings
  let selfopeningsRange = calendarSheet.getRange(weekRow + openingsValues.length + 1, weekCol, selfopeningsValues.length, nbCols);
  selfopeningsRange.setValues(selfopeningsValues);

  // -- Format & styles

  // Day
  let daysRange = calendarSheet.getRange(weekRow, weekCol + CALENDAR.DAY, currentRow - weekRow, 1);
  daysRange.setTextStyle(p.dayTextStyle);
  daysRange.setHorizontalAlignment("center");
  daysRange.setVerticalAlignment("middle");
  daysRange.setBackgrounds(daysBgs);
  // daysRange.setBorder(null, true, null, null, null, null, p.borderColor, null);

  // Hours
  let hoursRange = calendarSheet.getRange(weekRow, weekCol + CALENDAR.HOUR, currentRow - weekRow, 1);
  hoursRange.setTextStyle(p.hoursTextStyle);
  hoursRange.setHorizontalAlignment("center");
  hoursRange.setVerticalAlignment("middle");
  hoursRange.setBackgrounds(hoursBgs);

  // Slots
  let ceramistsSlotsRange = calendarSheet.getRange(weekRow, weekCol + CALENDAR.SLOT, currentRow - weekRow, p.slotsNames.length);
  ceramistsSlotsRange.setVerticalAlignment("middle");
  ceramistsSlotsRange.setBackground(null);
  ceramistsSlotsRange.setDataValidation(p.peopleRule);

  return titleRange.getNumRows() + currentRow - weekRow;
}
