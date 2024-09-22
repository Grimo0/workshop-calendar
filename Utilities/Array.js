
/**
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @return {Number[][]}
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
 * @return {String[]}
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