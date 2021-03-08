const ORIG_SHEET = '0';
const NEW_SHEET = '1';
const TIMEZONE = 'Asia/Tokyo';
const TIMEZONE_OFFSET = '0900';
const HOUR_MS = 1000 * 60 * 60;

function getTimeDiff(row) {
  try{
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ORIG_SHEET);
    const values = sheet.getRange('A' + String(row) + ':B' + String(row+1)).getValues();
    if (values[0][1] != 'entered') return '';
    return (getDateTime(values[1][0])-getDateTime(values[0][0]))/HOUR_MS;
  }catch(e){
    return '';
  }
}

function getDateTime(datetime){
  let ta = String(datetime).split(' ');
  ta.splice(3, 1);
  let time = ta[3].split(':');
  if(time[0] == '12')time[0] = '0';
  if(time[1].endsWith('PM'))time[0] = String(Number(time[0]) + 12);
  ta[3] = time[0] + ':' + time[1].replace('AM', '').replace('PM', '');
  ta.push('GMT+' + TIMEZONE_OFFSET);
  return new Date(ta.join(' '));
}

function getDayStart(datetime){
  let dt = new Date(datetime.getTime());
  dt.setHours(0, 0, 0, 0);
  return dt;
}

function makeEachDay(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TIMEZONE);
  const s_orig = ss.getSheetByName(ORIG_SHEET);

  const datetimes = s_orig.getRange('A:A').getValues();
  const inout = s_orig.getRange('B:B').getValues();
  let dt = null;
  let time = 0;
  let values = [];
  let i = -1;
  while(i < datetimes.length){
    i++;
    if(datetimes[i] == '')break;
    if(inout[i] != 'entered')continue;
    let dt_entered = getDateTime(datetimes[i]);

    i++;
    if(i == datetimes.length || datetimes[i] == '')break;
    let dt_exited = getDateTime(datetimes[i]);
    let dt_entered_start = getDayStart(dt_entered);
    let dt_exited_start = getDayStart(dt_exited);

    if(dt === null)dt = new Date(dt_entered_start.getTime());
    while(true){
      if(dt < dt_entered_start){
        values.push([new Date(dt.getTime()), time / HOUR_MS]);
        time = 0;
        dt.setDate(dt.getDate() + 1);
        continue;
      }
      if(dt.getTime() == dt_entered_start.getTime() && dt_entered_start.getTime() == dt_exited_start.getTime()){
        time += (dt_exited - dt_entered);
        break;        
      }
      if(dt.getTime() == dt_entered_start.getTime()){
        let dt_next = new Date(dt.getTime());
        dt_next.setDate(dt_next.getDate()+1);
        time += dt_next - dt_entered;
        values.push([new Date(dt.getTime()), time / HOUR_MS]);
        time = 0;
        dt.setDate(dt.getDate() + 1);
        continue;        
      }
      if(dt < dt_exited_start){
        values.push([new Date(dt.getTime()), 24]);
        time = 0;
        dt.setDate(dt.getDate() + 1);    
        continue;   
      }
      time = dt_exited - dt;
      break;
    }
  }
  if(time != 0){
    values.push([new Date(dt.getTime()), time / HOUR_MS]);
  }
  if(values.length == 0)return;

  let s_new = ss.getSheetByName(NEW_SHEET);
  if (!s_new) {
    s_new = ss.insertSheet(NEW_SHEET);
  }

  const range = s_new.getRange("A:B");
  range.deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  s_new.getRange(1, 1, 1, 2).setValues([['StartDateTime', 'Hours']]);
  s_new.getRange('A:A').setNumberFormat('yyyy/MM/dd');
  s_new.setFrozenRows(1);
  
  const numRows = values.length;
  const numColumns = values[0].length;
  s_new.getRange(2, 1, numRows, numColumns).setValues(values); 
}