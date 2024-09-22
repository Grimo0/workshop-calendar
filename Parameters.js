

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
   * @param {String} dayName Must contain the french name. Used to name the day but also to get the day index.
   * @param {Date} begin
   * @param {Date} end
   * @param {String} dayColor
   */
  constructor(dayName, begin, end, dayColor) {
    this.dayName = dayName;
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
}


class GenerateParameters {

  /**
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} activeSpreadsheet
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} publicSpreadsheet
   */
  constructor(activeSpreadsheet, publicSpreadsheet) {

    let peoplePublicSheet = publicSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);
    let peopleActiveSheet = activeSpreadsheet.getSheetByName(PEOPLE_SHEET_NAME);

    // -- Get Slots name
    let ceramistsSlotsName = activeSpreadsheet.getRangeByName('NomsColonnesTourneurs').getDisplayValues();
    this.ceramistsSlotsName = Array();
    for (let row = 0; row < ceramistsSlotsName.length; row++) {
      if (ceramistsSlotsName[row][0])
        this.ceramistsSlotsName.push(ceramistsSlotsName[row][0]);
    }
    let modelersSlotsName = activeSpreadsheet.getRangeByName('NomsColonnesModeleurs').getDisplayValues();
    this.modelersSlotsName = Array();
    for (let row = 0; row < modelersSlotsName.length; row++) {
      if (modelersSlotsName[row][0])
        this.modelersSlotsName.push(modelersSlotsName[row][0]);
    }
    let othersSlotsName = activeSpreadsheet.getRangeByName('NomsColonnesAutres').getDisplayValues();
    this.othersSlotsName = Array();
    for (let row = 0; row < othersSlotsName.length; row++) {
      if (othersSlotsName[row][0])
        this.othersSlotsName.push(othersSlotsName[row][0]);
    }

    this.slotsNames = this.ceramistsSlotsName.concat(this.modelersSlotsName, this.othersSlotsName);

    // -- Set the current date
    let daysToSkip = 7 - activeSpreadsheet.getRangeByName('JourFinDeSemaine').getValue();
    this.today = new Date();
    this.today.setDate(this.today.getDate() + (daysToSkip ? daysToSkip : 2));

    this.weeksToDisplay = activeSpreadsheet.getRangeByName('SemainesAffichees').getValue();

    // -- People data
    this.peopleNamesActiveRange = peopleActiveSheet.getRange(4, 1, peopleActiveSheet.getMaxRows());
    this.peopleNames = getFlatDisplayValues(this.peopleNamesActiveRange);

    this.ceramistsPastDaysActiveRange = peopleActiveSheet.getRange(4, 2, peopleActiveSheet.getMaxRows());
    this.ceramistsPastDays = getValuesAsNumber(this.ceramistsPastDaysActiveRange);

    this.ceramistsSelfPastDaysActiveRange = peopleActiveSheet.getRange(4, 6, peopleActiveSheet.getMaxRows());
    this.ceramistsSelfPastDays = getValuesAsNumber(this.ceramistsSelfPastDaysActiveRange);

    this.modelersPastDaysActiveRange = peopleActiveSheet.getRange(4, 10, peopleActiveSheet.getMaxRows());
    this.modelersPastDays = getValuesAsNumber(this.modelersPastDaysActiveRange);

    this.modelersSelfPastDaysActiveRange = peopleActiveSheet.getRange(4, 14, peopleActiveSheet.getMaxRows());
    this.modelersSelfPastDays = getValuesAsNumber(this.modelersSelfPastDaysActiveRange);

    // -- Dropdown validation rules
    let peopleNamePublicRange = peoplePublicSheet.getRange(2, 1, peoplePublicSheet.getMaxRows());
    this.peopleRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(peopleNamePublicRange)
      .setAllowInvalid(false)
      .build();

    // -- Special cells
    this.freeSlotCell = activeSpreadsheet.getRangeByName('EmplacementLibre').getCell(1, 1);
    this.unavailableSlotCell = activeSpreadsheet.getRangeByName('EmplacementIndisponible').getCell(1, 1);
    this.emptySlotCell = activeSpreadsheet.getRangeByName('EmplacementVide').getCell(1, 1);

    // -- Opening times
    let openingsRange = activeSpreadsheet.getRangeByName('PeriodesOuverture');
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

      this.openingTimes.push(new OpeningTime(openingRow[OPENING.DAY], begin, end, dayColor));
    }

    // -- Self-Opening times
    let selfopeningsRange = activeSpreadsheet.getRangeByName('PeriodesOuvertureLibre');
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

      this.selfopeningTimes.push(new OpeningTime(openingRow[OPENING.DAY], begin, end, dayColor));
    }

    // -- Closed times
    let closedRange = activeSpreadsheet.getRangeByName('PeriodesFermeture');
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
    this.borderColor = "#777777";

    this.headerTextStyle = SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setFontSize(12)
      .build();
    this.subheaderTextStyle = SpreadsheetApp.newTextStyle()
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
   * @param {String} peopleName
   * @param {number} slot
   * @param {number} days number of past days to add, default to 1.
   */
  addPastDay(peopleName, slot, days=1) {
    if (peopleName == "" || peopleName == this.freeSlotCell.getDisplayValue() || peopleName == this.unavailableSlotCell.getDisplayValue())
      return;

    let peopleNamesRows = this.peopleNamesActiveRange.getDisplayValues();
    for (let row = 0; row < peopleNamesRows.length; row++) {
      if (peopleNamesRows[row][0] == peopleName) {
        if (slot < this.ceramistsSlotsName.length) {
          this.ceramistsPastDays[row][0] += days;
          this.ceramistsSelfPastDays[row][0] += days;
        } else if (slot < this.ceramistsSlotsName.length + this.modelersSlotsName) {
          this.modelersPastDays[row][0] += days;
          this.modelersSelfPastDays[row][0] += days;
        }
        return;
      }
    }
  }
}