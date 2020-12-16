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
    return (getDate(end)-getDate(start))/1000/60/60/24;
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var values = sheet.getRange('A' + String(row) + ':B' + String(row+1)).getValues();
    if (values[0][1] != 'entered') return '';
    return (getDate(values[1][0])-getDate(values[0][0]))/1000/60/60/24;
  }catch(e){
    return '';
  }
}


function getDate(time){
  var ta = time.split(' ');
  ta.splice(3, 1);
  var time = ta[3].split(':');
  if(time[0] == '12')time[0] = '0';
  if(time[1].endsWith('PM'))time[0] = String(Number(time[0]) + 12);
  ta[3] = time[0] + ':' + time[1].replace('AM', '').replace('PM', '');
  return new Date(ta.join(' '));
}

function test(){
  var start = "September 9, 2020 at 09:44AM";
  var end = "September 9, 2020 at 09:11PM";
  getTimediffValues(start, end);
}

function test2(){
  getTimeDiff(2);
}

function test3(){
  getTimeDiff(3);
}