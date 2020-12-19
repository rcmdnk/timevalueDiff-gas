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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('0');
    var values = sheet.getRange('A' + String(row) + ':B' + String(row+1)).getValues();
    if (values[0][1] != 'entered') return '';
    return (getDate(values[1][0])-getDate(values[0][0]))/DAY_MS;
  }catch(e){
    return '';
  }
}


function getDate(date){
  var ta = String(date).split(' ');
  ta.splice(3, 1);
  var time = ta[3].split(':');
  if(time[0] == '12')time[0] = '0';
  if(time[1].endsWith('PM'))time[0] = String(Number(time[0]) + 12);
  ta[3] = time[0] + ':' + time[1].replace('AM', '').replace('PM', '');
  return new Date(ta.join(' '));
}


function getTimeOnDay(date){
  var d1 = getDate(date);
  var ta = String(date).split(' at ');
  var d0 = getDate(ta[0] + ' at 00:00:AM'); 
  return (d1 - d0) / DAY_MS;
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

function test4(){
  Logger.log('test');
}

function makeEachDay(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s_orig = ss.getSheetByName('0');

  var days = s_orig.getRange('A:A').getValues();
  var hours = s_orig.getRange('C:C').getValues();
  var values = [];
  days.forEach(function(d, i){
    if(hours[i] != ''){
      var date = getDate(d);
      var remain = hours[i];
      var remain_of_day = 1 - getTimeOnDay(d);
      while (true) {
        if(remain <= remain_of_day){
          values.push([new Date(date), remain * 24]);
          break
        }
        var time = remain_of_day;
        remain = remain - time;
        remain_of_day = 1;
        values.push([new Date(date), time * 24]);
        date.setDate(date.getDate() + 1);
        date.setHours(0);
        date.setMinutes(0);
        date.setSeconds(0);
        date.setMilliseconds(0);
      }
    }
  });
  var s_new = ss.getSheetByName('1');
  var range = s_new.getRange("A:B");
  range.deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  s_new.getRange(1, 1, 1, 2).setValues([['StartDateTime', 'Hours']]);
  s_new.getRange('A:A').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  s_new.setFrozenRows(1);
  
  var numRows = values.length;
  var numColumns = values[0].length;
  //sheet.insertRows(2,numRows);
  s_new.getRange(2, 1, numRows, numColumns).setValues(values); 
}