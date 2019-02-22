function doGet() {
  return HtmlService.createTemplateFromFile('index.html')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// get data from "Form Responese 1" for New Start Roster table by date
function getDataSortDate() {
  return SpreadsheetApp
    .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
    .getSheetByName('Form Responses 1')
    .getRange('A11:K')
    .sort(7)
    .getValues();
}
// get data for Current Roster Table by alphabet
function getDataSortName() {
  return SpreadsheetApp
    .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
    .getSheetByName('Form Responses 1')
    .getRange('A11:K')
    .sort(3)
    .getValues();
}
function getDataA_C() {
  return SpreadsheetApp
    .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
    .getSheetByName('Internal Dashboard A/C')
    .getRange(1, 1, 117, 14)
    .getValues();
}
function getDataB_D() {
  return SpreadsheetApp
    .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
    .getSheetByName('Internal Dashboard B/D')
    .getRange(1, 1, 117, 14)
    .getValues();
}
// determine current week is which crew
function detectCrew(now) {
  var first = new Date(now.getFullYear(), 0, 1);
  var weekNo = Math.ceil( (((now - first) / 86400000) + first.getDay() + 1) / 7 );
  if (now.getDay() == 0) {
    weekNo = weekNo - 1;
  }
  if (weekNo % 2 == 0) {
    return "A_C";
  } else {
    return "B_D";
  }
}
// get Monday in current week or next week based on param
function firstOfWeek(diff) {
  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var day = today.getDay();
  if (day == 0) {
      day = 7;
  }
  var monday = new Date(today.setDate(today.getDate() - day + diff));
  return monday;
}
// get Days on current week or next week based on param
function getWeek(diff) {
  var curr = new Date();
  var day = curr.getDay();
  if (day == 0) {
      day = 7;
  }
  var first = curr.getDate() - day + diff;
  var week = [];
  for (var i = 0; i < 7; i++) {
    var next = new Date(curr.getTime());
    next.setDate(first + i);
    next.setHours(0,0,0,0);
    week.push(next);
  }
  return week;
}
// determine if a person can be displayed or not
function filteredPerson(dept, crew, date, data) {
  var person = [];
  for (var i = 0; i < data.length; i++) {
    var start = "";
    var last = "";
    var formated_absence = [];
    if (data[i][6]) {
      start = new Date(new Date(data[i][6]).getTime() + 2 * 60 * 60 * 1000);
    }
    if (data[i][9]) {
      last = new Date(new Date(data[i][9]).getTime() + 2 * 60 * 60 * 1000);
    }
    if (data[i][10]) {        
      if (data[i][10].toString().indexOf(';') > -1) {
        formated_absence = data[i][10];
      } else {
        var absence = new Date(new Date(data[i][10]).getTime() + 2 * 60 * 60 * 1000);
        formated_absence = Utilities.formatDate(absence, "CST", "MM/dd/Y");
      }
    }
    if (start && data[i][3] == dept && data[i][4] == crew) {
      if (!last || last > date) {
        var formated_date = Utilities.formatDate(date, "CST", "MM/dd/Y");
        if (start < date) {
          if (formated_absence.indexOf(formated_date) > -1) {
            person.push(data[i][1] + " " + data[i][2] + '_A');  
          } else {
            person.push(data[i][1] + " " + data[i][2]);
          }
        } else if (start > date) {
          continue;
        } else {
          if (formated_absence && formated_absence.toString().indexOf(formated_date) > -1) {
            person.push(data[i][1] + " " + data[i][2] + '_B');
          } else {
            person.push(data[i][1] + " " + data[i][2] + '_S');
          }
        }
      }
    }
  }
  return person;
}
function makePersons(crew, dept, week, data, order) {
  var persons = [];
  if (crew == "A_C") {
    if (order == "first") {
      if (dept == 'Roll Fed' || 'Inline') {
        var crews = ["A Crew", "A Crew", "B Crew", "B Crew", "A Crew", "A Crew", "A Crew"];
      }
      if (dept == 'Eco Star') {
        var crews = ["Days", "Days", "Days", "Days", "OFF", "OFF", "OFF"];
      }
      if (dept == 'North Plant') {
        var crews = ["Days", "Days", "Days", "Days", "OFF", "OFF", "OFF"];
      }
    } else {
      if (dept == 'Roll Fed' || 'Inline') {
        var crews = ["C Crew", "C Crew", "D Crew", "D Crew", "C Crew", "C Crew", "C Crew"];
      }
      if (dept == 'Eco Star') {
        var crews = ["Nights", "Nights", "Nights", "Nights", "OFF", "OFF", "OFF"];
      }
      if (dept == 'North Plant') {
        var crews = ["Nights", "Nights", "Nights", "Nights", "OFF", "OFF", "OFF"];
      }
    }
  }
  if (crew == "B_D") {
    if (order == "first") {
      if (dept == 'Roll Fed' || 'Inline') {
        var crews = ["B Crew", "B Crew", "A Crew", "A Crew", "B Crew", "B Crew", "B Crew"];
      }
      if (dept == 'Eco Star') {
        var crews = ["Days", "Days", "Days", "OFF", "OFF", "OFF", "OFF"];
      }
      if (dept == 'North Plant') {
        var crews = ["Days", "Days", "Days", "Days", "OFF", "OFF", "OFF"];
      }
    } else {
      if (dept == 'Roll Fed' || 'Inline') {
        var crews = ["D Crew", "D Crew", "C Crew", "C Crew", "D Crew", "D Crew", "D Crew"];
      }
      if (dept == 'Eco Star') {
        var crews = ["Nights", "Nights", "Nights", "OFF", "OFF", "OFF", "OFF"];
      }
      if (dept == 'North Plant') {
        var crews = ["Nights", "Nights", "Nights", "Nights", "OFF", "OFF", "OFF"];
      }
    }
  }
  for (var i = 0; i < 7; i++) {
    persons[i] = filteredPerson(dept, crews[i], week[i], data);
  }
  return persons;
}
function insertTitle(activeSheet, color, titleName) {
  activeSheet.getRange('A2').setBackground(color);
  activeSheet.getRange('A2').setValue(titleName);
}
function insertDate(activeSheet, row, data) {
  for (var i in row) {
    activeSheet.getRange('A'+row[i]).setValue([data[0]]);
    activeSheet.getRange('C'+row[i]).setValue([data[1]]);
    activeSheet.getRange('E'+row[i]).setValue([data[2]]);
    activeSheet.getRange('G'+row[i]).setValue([data[3]]);
    activeSheet.getRange('I'+row[i]).setValue([data[4]]);
    activeSheet.getRange('K'+row[i]).setValue([data[5]]);
    activeSheet.getRange('M'+row[i]).setValue([data[6]]);
  }
}
function insertData(activeSheet, ref, data) {
  var row;
  switch (ref) {
    case 'firstRollA_C':
      row = 9;
      break;
    case 'secondRollA_C':
      row = 21;
      break;
    case 'firstInA_C':
      row = 38;
      break;
    case 'secondInA_C':
      row = 50;
      break;
    case 'firstEcoA_C':
      row = 67;
      break;
    case 'secondEcoA_C':
      row = 79;
      break;
    case 'firstNorthA_C':
      row = 96;
      break;
    case 'secondNorthA_C':
      row = 108;
      break;
    
    case 'firstRollB_D':
      row = 9;
      break;
    case 'secondRollB_D':
      row = 21;
      break;
    case 'firstInB_D':
      row = 38;
      break;
    case 'secondInB_D':
      row = 50;
      break;
    case 'firstEcoB_D':
      row = 67;
      break;
    case 'secondEcoB_D':
      row = 79;
      break;
    case 'firstNorthB_D':
      row = 96;
      break;
    case 'secondNorthB_D':
      row = 108;
      break;
  }
  for (var i = 0; i < 7; i++) {
    for (var j = 0; j < data[i].length; j++) {
      activeSheet.getRange(row+j, 2*i+1).setValue(data[i][j]);
    }
  }
}
// it should be run once per week
function insertDataToSheet() {
  var spreadSheetId = '1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA';
  var sheetNameA_C = 'Internal Dashboard A/C';
  var sheetNameB_D = 'Internal Dashboard B/D';
  var currentWeek = getWeek(1);
  var nextWeek = getWeek(8);
  var currentFormatDate = [];
  var nextFormatDate = [];
  for (var i in currentWeek) {
    currentFormatDate.push(Utilities.formatDate(currentWeek[i], "CST", "MM/dd/Y"));
  }
  for (var i in nextWeek) {
    nextFormatDate.push(Utilities.formatDate(nextWeek[i], "CST", "MM/dd/Y"));
  }
  var crew = detectCrew(new Date());
  if (crew == 'A_C') {
    var colorA_C = '#00FF00'; //green
    var colorB_D = '#FF0000'; //red

    var titleA_C = 'CURRENT';
    var titleB_D = 'NEXT';

    var formatDateA_C = currentFormatDate;
    var formatDateB_D = nextFormatDate;

    var weekA_C = currentWeek;
    var weekB_D = nextWeek;
  }
  if (crew == 'B_D') {
    var colorA_C = '#FF0000'; //red
    var colorB_D = '#00FF00'; //green

    var titleA_C = 'NEXT';
    var titleB_D = 'CURRENT';

    var formatDateA_C = nextFormatDate;
    var formatDateB_D = currentFormatDate;

    var weekA_C = nextWeek;
    var weekB_D = currentWeek;
  }
  var sourceData = getDataSortName();
  // A/C Roll Fed
  var firstRollA_C = makePersons('A_C', 'Roll Fed', weekA_C, sourceData, 'first');
  var secondRollA_C = makePersons('A_C', 'Roll Fed', weekA_C, sourceData, 'second');
  // A/C Inline
  var firstInA_C = makePersons('A_C', 'Inline', weekA_C, sourceData, 'first');
  var secondInA_C = makePersons('A_C', 'Inline', weekA_C, sourceData, 'second');
  // A/C Eco Star
  var firstEcoA_C = makePersons('A_C', 'Eco Star', weekA_C, sourceData, 'first');
  var secondEcoA_C = makePersons('A_C', 'Eco Star', weekA_C, sourceData, 'second');
  // A/C North Plant
  var firstNorthA_C = makePersons('A_C', 'North Plant', weekA_C, sourceData, 'first');
  var secondNorthA_C = makePersons('A_C', 'North Plant', weekA_C, sourceData, 'second');

  // B/D Roll Fed
  var firstRollB_D = makePersons('B_D', 'Roll Fed', weekB_D, sourceData, 'first');
  var secondRollB_D = makePersons('B_D', 'Roll Fed', weekB_D, sourceData, 'second');
  // B/D Inline
  var firstInB_D = makePersons('B_D', 'Inline', weekB_D, sourceData, 'first');
  var secondInB_D = makePersons('B_D', 'Inline', weekB_D, sourceData, 'second');
  // B/D Eco Star
  var firstEcoB_D = makePersons('B_D', 'Eco Star', weekB_D, sourceData, 'first');
  var secondEcoB_D = makePersons('B_D', 'Eco Star', weekB_D, sourceData, 'second');
  // B/D North Plant
  var firstNorthB_D = makePersons('B_D', 'North Plant', weekB_D, sourceData, 'first');
  var secondNorthB_D = makePersons('B_D', 'North Plant', weekB_D, sourceData, 'second');
  
  var activeSpreadsheet = SpreadsheetApp.openById(spreadSheetId);
  var activeSheetA_C = activeSpreadsheet.getSheetByName(sheetNameA_C);
  var activeSheetB_D = activeSpreadsheet.getSheetByName(sheetNameB_D);
  // insert title/date/color to A/C
  insertTitle(activeSheetA_C, colorA_C, titleA_C);
  insertDate(activeSheetA_C, [6, 35, 64, 93], formatDateA_C);
  // insert data to A/C
  insertData(activeSheetA_C, 'firstRollA_C', firstRollA_C);
  insertData(activeSheetA_C, 'secondRollA_C', secondRollA_C);
  insertData(activeSheetA_C, 'firstInA_C', firstInA_C);
  insertData(activeSheetA_C, 'secondInA_C', secondInA_C);
  insertData(activeSheetA_C, 'firstEcoA_C', firstEcoA_C);
  insertData(activeSheetA_C, 'secondEcoA_C', secondEcoA_C);
  insertData(activeSheetA_C, 'firstNorthA_C', firstNorthA_C);
  insertData(activeSheetA_C, 'secondNorthA_C', secondNorthA_C);

  // insert title/date/color to B/D
  insertTitle(activeSheetB_D, colorB_D, titleB_D);
  insertDate(activeSheetB_D, [6, 35, 64, 93], formatDateB_D);
  // insert data to B/D
  insertData(activeSheetB_D, 'firstRollB_D', firstRollB_D);
  insertData(activeSheetB_D, 'secondRollB_D', secondRollB_D);
  insertData(activeSheetB_D, 'firstInB_D', firstInB_D);
  insertData(activeSheetB_D, 'secondInB_D', secondInB_D);
  insertData(activeSheetB_D, 'firstEcoB_D', firstEcoB_D);
  insertData(activeSheetB_D, 'secondEcoB_D', secondEcoB_D);
  insertData(activeSheetB_D, 'firstNorthB_D', firstNorthB_D);
  insertData(activeSheetB_D, 'secondNorthB_D', secondNorthB_D);
}
function getMaxLength(data, ref) {
  var row;
  var maxLength;
  switch (ref) {
    case 'rollFirst':
      row = 8;
      break;
    case 'rollSecond':
      row = 20;
      break;
    case 'inFirst':
      row = 37;
      break;
    case 'inSecond':
      row = 49;
      break;
    case 'ecoFirst':
      row = 66;
      break;
    case 'ecoSecond':
      row = 78;
      break;
    case 'northFirst':
      row = 95;
      break;
    case 'northSecond':
      row = 107;
      break;
  }
  maxLength = Math.max(data[row-1][1], data[row-1][3], data[row-1][5], data[row-1][7], data[row-1][9], data[row-1][11], data[row-1][13]);
  return [row, maxLength];
}
function makeClassName(i, j, ref, data){
  var row_length = getMaxLength(data, ref);
  var rowNum = row_length[0];
  var length = data[rowNum-1][2*j+1];
  if (i >= length) {
    return 'outside';
  } else {
    return 'normal';
  }
}
function makeTableFromInternal(data, ref) {
  // get maxLength
  var row_length = getMaxLength(data, ref);
  var rowNum = row_length[0];
  var maxLength = row_length[1];
  
  // define table maxLength * 7
  var table = [];
  for (var i = 0; i < maxLength; i++) {
    table[i] = [];
  }
  // insert data from sheet to table
  for (var i = 0; i < maxLength; i++) {
    for (var j = 0; j < 7; j++) {
      if (data[i+rowNum][2*j+1] && data[i+rowNum][2*j+1] != '-') {
        table[i][j] = data[i+rowNum][2*j+1];
      } else {
        table[i][j] = null;
      }
    }
  }
  // remove null in table
  var columns = [];
  for (var i = 0; i < 7; i++) {
    columns[i] = [];
    for (j = 0; j < table.length; j++) {
      if (table[j][i]) columns[i].push(table[j][i]);
    }
  }
  var new_table = [];
  for (var i = 0; i < table.length; i++) {
    new_table[i] = [];
    for (j = 0; j < 7; j++) {
      if (columns[j][i]) new_table[i][j] = columns[j][i];
      else new_table[i][j] = null;
    }
  }
  return new_table;
}