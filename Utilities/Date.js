
/**
 * @param {number} weekNumber
 * @param {number} year
 * @returns {Date} The date of the monday in the given week
 */
function getDateOfWeek(weekNumber, year) {
  var date = new Date(year, 0, 1 + (weekNumber - 1) * 7); // 1st of January + 7 days for each week
  // The first week must contain the first Thursday (4th day)
  var dow = date.getDay();
  if (dow <= 4)
    date.setDate(date.getDate() - dow + 1); // Remove the missing days of the week from the first week
  else
    date.setDate(date.getDate() - dow + 8); // Add the number of days in the year but prior to the first week
  return date;
}


/**
 * @param {Date} d
 * @returns {number} The number of the week considering that they starts on monday
 * and the first week of the year contains the first Thursday
 */
function getWeekNumber(d) {
  // Copy date so don't modify original
  d = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  // Set to nearest Thursday: current date + 4 - current day number
  // Make Sunday's day number 7
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  // Get first day of year
  var yearStart = new Date(d.getFullYear(), 0, 1);
  // Calculate full weeks to nearest Thursday
  var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  // Return array of year and week number
  return weekNo;
}


/**
 * @param {String} s Must be the name of the day in french (case don't matter).
 * @returns {number} The 0 indexed placement of the day in the week or -1 if the name is invalid
 */
function getDayIdx(s) {
  var s = s.toUpperCase();
  if (s.includes("DIMANCHE"))
    return 6;
  if (s.includes("SAMEDI"))
    return 5;
  if (s.includes("VENDREDI"))
    return 4;
  if (s.includes("JEUDI"))
    return 3;
  if (s.includes("MERCREDI"))
    return 2;
  if (s.includes("MARDI"))
    return 1;
  if (s.includes("LUNDI"))
    return 0;
  return -1;
}

/**
 * Update a date based on a formatted date string.
 * @param {Date} d The date to update
 * @param {String} str A french formatted date "DD[/MM[/YY[YY]]]"". All the missing elements are taken from the current date and time.
 */
function updateDate(d, str) {
  let daySplit = str.split("/");
  if (daySplit.length > 2)
    d.setFullYear(daySplit[2] < 2000 ? 2000 + parseInt(daySplit[2]) : daySplit[2], daySplit[1] - 1, daySplit[0]);
  else if (daySplit.length > 1)
    d.setMonth(daySplit[1] - 1, daySplit[0]);
  else
    d.setDate(daySplit[0]);
}

/**
 * Update a date based on a formatted time string.
 * @param {Date} d The date to update
 * @param {String} str An hour in the format "HH[h[MM]]". All the missing elements are taken from the current date and time.
 */
function updateTime(d, str) {
  let hourSplit = str.split("h");
  if (hourSplit.length > 1)
    d.setHours(hourSplit[0], hourSplit[1], 0);
  else
    d.setHours(hourSplit[0], 0, 0);
}