const ORIG_SHEET = '0';
const NEW_SHEET = '1';
const TIMEZONE = 'Asia/Tokyo';
const TIMEZONE_OFFSET = '0900';
const DAY_MS = 1000 * 60 * 60 * 24;

/**
 * Fill timevalue difference of end -start, if is_exited is 'exited', else empty
 * @param {'exited'} is_exited exited or entered
 * @param {'September 9, 2020 at 08:44AM'} start start datetime
 * @param {'September 9, 2020 at 07:11PM'} end end datetime
 * @return timevalue difference or '' if is_exited is 'entered'
 * @customfunction
 */
function getTimevalueDiffAtExited(is_exited, start, end) {
  if(is_exited != 'exited')return '';
  try{
    return (getDate(end)-getDate(start))/DAY_MS;
  }catch(e){
    return '';
  }
}

/**
 * Fill timevalue difference of `A{row}` - `A{row-1}`, if `B{row}` is 'exited', else empty
 * @param {'4'} row use `row()` to use current row
 * @return timevalue difference or '' if is_exited is 'entered'
 * @customfunction
 */
function getTimeDiff(row) {
  try{
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ORIG_SHEET);
    const values = sheet.getRange('A' + String(row) + ':B' + String(row+1)).getValues();
    if (values[0][1] != 'entered') return '';
    return (getDate(values[1][0])-getDate(values[0][0]))/DAY_MS;
  }catch(e){
    return '';
  }
}


function getDate(date){
  let ta = String(date).split(' ');
  ta.splice(3, 1);
  let time = ta[3].split(':');
  if(time[0] == '12')time[0] = '0';
  if(time[1].endsWith('PM'))time[0] = String(Number(time[0]) + 12);
  ta[3] = time[0] + ':' + time[1].replace('AM', '').replace('PM', '');
  ta.push('GMT+' + TIMEZONE_OFFSET);
  return new Date(ta.join(' '));
}

function getTimeOnDay(date){
  const d1 = getDate(date);
  const ta = String(date).split(' at ');
  const d0 = getDate(ta[0] + ' at 00:00:AM'); 
  return (d1 - d0) / DAY_MS;
}

function getDayStart(date){
  let day_start = new Date(date);
  day_start.setHours(0);
  day_start.setMinutes(0);
  day_start.setSeconds(0);
  day_start.setMilliseconds(0);
  return day_start;
}

function makeEachDay(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TIMEZONE);
  const s_orig = ss.getSheetByName(ORIG_SHEET);

  let days = s_orig.getRange('A:A').getValues();
  let hours = s_orig.getRange('C:C').getValues();
  let prev = '';
  let values = [];
  days.forEach(function(d, i){
    if(hours[i] != ''){
      let date = getDate(d);
      let remain = hours[i];
      let remain_of_day = 1 - getTimeOnDay(d);
      while (true) {
        if(prev != ''){
          while(date - prev > DAY_MS){
            prev.setDate(prev.getDate() + 1);
            values.push([new Date(prev), 0]);
          }
        }
        if(remain <= remain_of_day){
          values.push([new Date(date), remain * 24]);
          prev = getDayStart(date);
          break
        }
        const time = remain_of_day;
        remain = remain - time;
        remain_of_day = 1;
        values.push([new Date(date), time * 24]);
        prev = getDayStart(date);
        date.setDate(date.getDate() + 1);
        date = getDayStart(date);
      }
    }
  });
  const s_new = ss.getSheetByName(NEW_SHEET);
  const range = s_new.getRange("A:B");
  range.deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  s_new.getRange(1, 1, 1, 2).setValues([['StartDateTime', 'Hours']]);
  s_new.getRange('A:A').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  s_new.setFrozenRows(1);
  
  const numRows = values.length;
  const numColumns = values[0].length;
  s_new.getRange(2, 1, numRows, numColumns).setValues(values); 
}