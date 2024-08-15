const dayjs = require('dayjs');
const utc = require('dayjs/plugin/utc');
const timezone = require('dayjs/plugin/timezone');

// Config date
dayjs.extend(utc)
dayjs.extend(timezone);

function convertMsToDate(ms) {
  let time = dayjs.tz(ms, "Asia/Ho_Chi_Minh");
  let date = time.add(time.utcOffset(), 'minutes').toDate();
  return date;
}

function convertISO8601StrToDate(str, addingHour=0) {
  let time = dayjs(str);
  if (addingHour != undefined) {time = time.add(addingHour, 'hour');}
  let date = time.add(time.utcOffset(), 'minutes').toDate();
  return date;
}

module.exports = {convertMsToDate, convertISO8601StrToDate};