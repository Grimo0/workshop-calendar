
/**
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @returns {Number[][]}
 */
function getValuesAsNumber(range) {
	/** @type {Number[][]} */
	let rv = Array();
	let rows = range.getDisplayValues();
	for (let row of rows) {
		let converted_row = Array();
		for (let cell of row) {
			let value = parseInt(cell);
			if (isNaN(value))
				value = 0;
			converted_row.push(value);
		}
		rv.push(converted_row)
	}
	return rv
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @returns {String[]}
 */
function getFlatDisplayValues(range) {
	/** @type {String[]} */
	let rv = Array();
	let rows = range.getDisplayValues();
	for (let row of rows) {
		let converted_row = Array();
		for (let cell of row) {
			if (cell && cell.length > 0)
				converted_row.push(cell);
		}
		if (converted_row.length > 0)
			rv.push(converted_row)
	}
	return rv
}

/**
 * @param {String[][]} values
 * @returns {[Map<[Date, Date], Array<String>>, Map<[Date, Date], Array<String>>]} [daysMap, selfDaysMap]
 */
function getDaysMap(values) {
  /** @type {Map<[Date, Date], Array<String>>} */
  let daysMap = new Map();
  /** @type {Map<[Date, Date], Array<String>>} */
  let selfDaysMap = new Map();
  let isSelfDay = false;
  for (let r = 0; r < values.length; r++) {

    // Update isSelfDay and skip empty rows
    if (values[r][CALENDAR.HOUR] == "") {
      isSelfDay = (values[r][0] == SELF_DAYS_HEADER);
      continue;
    }

    let row = values[r];
    let beginDate = new Date();

    // Get begin/end date
    let daySplit = row[CALENDAR.DAY].trim().split(" ");
    if (daySplit.length > 1) {
      updateDate(beginDate, daySplit[daySplit.length - 1]);
    } else {
      continue;
    }

    let endDate = new Date();
    let hourSplit = row[CALENDAR.HOUR].split("-");
    if (hourSplit.length > 1) {
      updateTime(beginDate, hourSplit[0]);

      endDate.setFullYear(beginDate.getFullYear(), beginDate.getMonth(), beginDate.getDate());
      updateTime(endDate, hourSplit[1]);
    } else {
      continue;
    }

    if (isSelfDay) {
      selfDaysMap.set([beginDate, endDate], row);
    } else {
      daysMap.set([beginDate, endDate], row);
    }
  }

  return [daysMap, selfDaysMap]
}