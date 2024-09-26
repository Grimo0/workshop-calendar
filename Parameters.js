

const OPENING_TYPE = class {
  static get REGULAR() { return "Encadré"; }
  static get SELF() { return "Libre"; }
}
const CALENDAR = class {
  static get TYPE() { return 0; }
  static get DAY() { return 1; }
  static get HOUR() { return 2; }
  static get SLOT() { return 3; }
}
const OPENING = class {
  static get DAY() { return 0; }
  static get BEGIN() { return 1; }
  static get END() { return 2; }
}
const CLOSED = class {
  static get BEGIN_DAY() { return 0; }
  static get BEGIN_HOUR() { return 1; }
  static get END_DAY() { return 2; }
  static get END_HOUR() { return 3; }
}


class ClosedTime {
  /**
   * @param {Date} begin
   * @param {Date} end
   */
  constructor(begin, end) {
    this.begin = begin;
    this.end = end;
  }
}


class OpeningTime {
  /**
   * @param {String} type
   * @param {String} dayName Must contain the french name. Used to name the day but also to get the day index.
   * @param {Date} begin
   * @param {Date} end
   * @param {String} dayColor
   */
  constructor(type, dayName, begin, end, dayColor) {
    this.type = type;
    this.dayName = dayName.slice(0, 1).toUpperCase() + dayName.slice(1, 3);
    this.dayOfWeek = getDayIdx(dayName);
    this.begin = begin;
    this.end = end;
    this.dayColor = dayColor;

    let dayColorInt = parseInt(dayColor.slice(1), 16);
    let r = (dayColorInt & 0xff0000) >> 16;
    let g = (dayColorInt & 0xff00) >> 8;
    let b = dayColorInt & 0xff;
    r += Math.round((255 - r) / 3);
    g += Math.round((255 - g) / 3);
    b += Math.round((255 - b) / 3);
    this.hourColor = "#" + r.toString(16) + g.toString(16) + b.toString(16);
  }


  /**
   * @param {Date} dateOfWeek
   * @returns {Date}
   */
  getBeginInWeek(dateOfWeek) {
    return new Date(
      dateOfWeek.getFullYear(),
      dateOfWeek.getMonth(),
      dateOfWeek.getDate() + this.dayOfWeek,
      this.begin.getHours(),
      this.begin.getMinutes()
    );
  }


  /**
   * @param {Date} dateOfWeek
   * @returns {Date}
   */
  getEndInWeek(dateOfWeek) {
    return new Date(
      dateOfWeek.getFullYear(),
      dateOfWeek.getMonth(),
      dateOfWeek.getDate() + this.dayOfWeek,
      this.end.getHours(),
      this.end.getMinutes()
    );
  }
}


class GenerateParameters {

  /**
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} activeSpreadsheet
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} publicSpreadsheet
   */
  constructor(activeSpreadsheet, publicSpreadsheet) {

    let peoplePublicSheet = publicSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
    let peopleActiveSheet = activeSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
    let parametersSheet = activeSpreadsheet.getSheetByName(PARAMETERS_SHEET_NAME);

    // -- Get Slots name
    let categoriesSheet = activeSpreadsheet.getSheetByName(CATEGORIES_SHEET_NAME);
    this.ceramistsSlotsName = getFlatDisplayValues(categoriesSheet.getRange(2, 2, 1, categoriesSheet.getMaxColumns() - 1));
    this.modelersSlotsName = getFlatDisplayValues(categoriesSheet.getRange(3, 2, 1, categoriesSheet.getMaxColumns() - 1));
    this.othersSlotsName = getFlatDisplayValues(categoriesSheet.getRange(4, 2, 1, categoriesSheet.getMaxColumns() - 1));

    this.slotsNames = this.ceramistsSlotsName.concat(this.modelersSlotsName, this.othersSlotsName);

    // -- Set the current date
    let daysToSkip = 7 - parametersSheet.getRange(2, 2).getValue();
    this.today = new Date();
    this.today.setDate(this.today.getDate() + (daysToSkip ? daysToSkip : 2));

    this.weeksToDisplay = parametersSheet.getRange(1, 2).getValue();

    // -- People data
    // Names
    this.peopleNamesActiveRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 1, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
    this.peopleNames = getFlatDisplayValues(this.peopleNamesActiveRange);

    // Past days
    this.ceramistsPastDaysActiveRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 2, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
    this.ceramistsPastDays = getValuesAsNumber(this.ceramistsPastDaysActiveRange);

    this.ceramistsSelfPastDaysActiveRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 6, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
    this.ceramistsSelfPastDays = getValuesAsNumber(this.ceramistsSelfPastDaysActiveRange);

    this.modelersPastDaysActiveRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 10, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
    this.modelersPastDays = getValuesAsNumber(this.modelersPastDaysActiveRange);

    this.modelersSelfPastDaysActiveRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 14, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);
    this.modelersSelfPastDays = getValuesAsNumber(this.modelersSelfPastDaysActiveRange);

    // Future days
    this.ceramistsDaysToComeActiveRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 3, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);

    this.ceramistsSelfDaysToComeActiveRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 7, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);

    this.modelersDaysToComeActiveRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 11, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);

    this.modelersSelfDaysToComeActiveRange = peopleActiveSheet.getRange(PEOPLE_HEADER_NB_ROWS + 1, 15, peopleActiveSheet.getMaxRows() - PEOPLE_HEADER_NB_ROWS);

    // -- Dropdown validation rules
    let peopleNamePublicRange = peoplePublicSheet.getRange(2, 1, peoplePublicSheet.getMaxRows() - 1);
    this.peopleRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(peopleNamePublicRange)
      .setAllowInvalid(false)
      .build();

    // -- Special cells
    this.freeSlotCell = parametersSheet.getRange(4, 2).getCell(1, 1);
    this.unavailableSlotCell = parametersSheet.getRange(5, 2).getCell(1, 1);
    this.emptySlotCell = parametersSheet.getRange(6, 2).getCell(1, 1);

    // -- Opening times
    let openingsSheet = activeSpreadsheet.getSheetByName(OPENINGS_SHEET_NAME);
    let openingsRange = openingsSheet.getRange(3, 1, openingsSheet.getMaxRows() - 2, 3);
    let openingsValues = openingsRange.getDisplayValues();
    /** @type {OpeningTime[]} */
    this.openingTimes = Array();
    for (let row = 0; row < openingsValues.length; row++) {
      let openingRow = openingsValues[row];
      if (openingRow[0].length == 0)
        break;

      let begin = new Date();
      updateTime(begin, openingRow[OPENING.BEGIN]);

      let end = new Date();
      updateTime(end, openingRow[OPENING.END]);

      let dayColor = openingsRange.getCell(1 + row, 1 + OPENING.DAY).getBackground();

      this.openingTimes.push(new OpeningTime(OPENING_TYPE.REGULAR, openingRow[OPENING.DAY], begin, end, dayColor));
    }

    // -- Self-Opening times
    let selfopeningsRange = openingsSheet.getRange(3, 4, openingsSheet.getMaxRows() - 2, 3);
    let selfopeningsValues = selfopeningsRange.getDisplayValues();
    /** @type {OpeningTime[]} */
    this.selfopeningTimes = Array();
    for (let row = 0; row < selfopeningsValues.length; row++) {
      let openingRow = selfopeningsValues[row];
      if (openingRow[0].length == 0)
        break;

      let begin = new Date();
      updateTime(begin, openingRow[OPENING.BEGIN]);

      let end = new Date();
      updateTime(end, openingRow[OPENING.END]);

      let dayColor = selfopeningsRange.getCell(1 + row, 1 + OPENING.DAY).getBackground();

      this.selfopeningTimes.push(new OpeningTime(OPENING_TYPE.SELF, openingRow[OPENING.DAY], begin, end, dayColor));
    }

    // -- Closed times
    let closedSheet = activeSpreadsheet.getSheetByName(CLOSED_SHEET_NAME);
    let closedRange = closedSheet.getRange(3, 1, closedSheet.getMaxRows() - 2, 4);
    let closedValues = closedRange.getDisplayValues();
    /** @type {ClosedTime[]} */
    this.closedTimes = Array();
    for (let closedRow of closedValues) {
      if (closedRow[0].length == 0)
        break;
      let begin = new Date()
      updateDate(begin, closedRow[CLOSED.BEGIN_DAY]);
      updateTime(begin, closedRow[CLOSED.BEGIN_HOUR]);

      let end = new Date();
      if (closedRow[CLOSED.END_DAY].length == 0) {
        end.setFullYear(begin.getFullYear(), begin.getMonth(), begin.getDate() + 1);
        end.setHours(0, 0, 0);
      } else {
        updateDate(end, closedRow[CLOSED.END_DAY]);

        if (closedRow[CLOSED.END_HOUR] == "") {
          end.setHours(0, 0, 0);
          end.setDate(end.getDate() + 1);
        } else {
          updateTime(end, closedRow[CLOSED.END_HOUR]);
        }
      }

      this.closedTimes.push(new ClosedTime(begin, end));
    }

    // -- Styles
    this.headerBackground = "#d9d9d9";
    this.subheaderBackground = "#e9e9e9";
    this.separatorBackground = "#e9e9e9";
    this.borderColor = "#777777";

    this.headerTextStyle = SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setFontSize(12)
      .build();
    this.subheaderTextStyle = SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setFontSize(12)
      .build();
    this.separatorTextStyle = SpreadsheetApp.newTextStyle()
      .setFontSize(11)
      .build();
    this.dayTextStyle = SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setFontSize(11)
      .build();
    this.hoursTextStyle = SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setFontSize(11)
      .build();
  }


  /**
   * @param {Date} begin
   * @param {Date} end
   * @returns {boolean}
   */
  isClosed(begin, end) {
    for (let closedTime of this.closedTimes) {
      if (end.getTime() >= closedTime.begin.getTime()
        && begin.getTime() <= closedTime.end.getTime()) {
        return true;
      }
    }
    return false;
  }


  /**
   * @param {String} peopleName
   * @param {number} slot
   * @param {boolean} self
   * @param {number} days number of past days to add, default to 1.
   */
  addPastDay(peopleName, slot, self, days = 1) {
    if (peopleName == "" || peopleName == this.freeSlotCell.getDisplayValue() || peopleName == this.unavailableSlotCell.getDisplayValue())
      return;

    let peopleNamesRows = this.peopleNamesActiveRange.getDisplayValues();
    for (let row = 0; row < peopleNamesRows.length; row++) {
      if (peopleNamesRows[row][0] == peopleName) {
        if (slot < this.ceramistsSlotsName.length) {
          if (self)
            this.ceramistsSelfPastDays[row][0] += days;
          else
            this.ceramistsPastDays[row][0] += days;
        } else if (slot < this.ceramistsSlotsName.length + this.modelersSlotsName.length) {
          if (self)
            this.modelersSelfPastDays[row][0] += days;
          else
            this.modelersPastDays[row][0] += days;
        }
        return;
      }
    }
  }
}