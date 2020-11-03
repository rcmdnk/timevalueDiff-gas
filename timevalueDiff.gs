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
    var values = sheet.getRange('A' + String(row-1) + ':B' + String(row)).getValues();
    if (values[1][1] != 'exited') return '';
    return (getDate(values[1][0])-getDate(values[0][0]))/1000/60/60/24;
  }catch(e){
    return '';
  }
}


function getDate(time){
  var ta = time.split(' ');
  ta.splice(3, 1);
  if (ta[3].endsWith('PM')){
    var time = ta[3].replace('PM', '').split(':');
    ta[3] = String(Number(time[0]) + 12) + ':' + time[1];
  }else{
    ta[3] = ta[3].replace('AM', '');
  }
  return new Date(ta.join(' '));
}

function test(){
  var start = "September 9, 2020 at 09:44AM";
  var end = "September 9, 2020 at 09:11PM";
  getTimediffValues(start, end);
}

function test2(){
  getTimediff(2);
}

function test3(){
  getTimediff(3);
}