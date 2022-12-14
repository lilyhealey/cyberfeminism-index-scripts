/**
 * Returns the current time as a formatted string.
 * 
 * e.g. 'Jan-01-1970-at-0000'
 * 
 * @returns {String}
 */
 function timestamp() {

  var months = [
    'jan', 'feb', 'mar',
    'apr', 'may', 'jun',
    'jul', 'aug', 'sep',
    'oct', 'nov', 'dec',
  ];
  var now = new Date();
  var hour = now.getHours();
  var minute = now.getMinutes();
  var month = now.getMonth();
  var day = now.getDate();
  var year = now.getFullYear();

  // 15
  hour = hour.toString();
  hour = hour.padStart(2, '0');

  // e.g 47
  minute = minute.toString();
  minute = minute.padStart(2, '0');

  // jan
  month = months[month];

  // 01
  day = day.toString();
  day = day.padStart(2, '0');

  // 2020
  year = year.toString();

  // 1547
  var t = [hour, minute].join('');

  // 01-jan-2020
  var d = [day, month, year].join('-');

  // jan-01-2020-at-1547
  var dt = [d, t].join('-at-');

  return dt;
}
