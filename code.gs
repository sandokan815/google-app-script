  function doGet() {
    return HtmlService.createTemplateFromFile('index.html')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  function getDataSortDate() {
    return SpreadsheetApp
      .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
      .getSheetByName('Form Responses 1')
      .getRange('A11:L')
      .sort(7)
      .getValues();
  }
  function getDataSortName() {
    return SpreadsheetApp
      .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
      .getSheetByName('Form Responses 1')
      .getRange('A11:L')
      .sort(3)
      .getValues();
  }
  function getFirstSchedule() {
    return SpreadsheetApp
      .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
      .getSheetByName('Daily Schedule')
      .getRange(4, 2, 28, 10)
      .getValues();
  }
  function getSecondSchedule() {
    return SpreadsheetApp
      .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
      .getSheetByName('Eco, North & Warehouse Schedule')
      .getRange(4, 2, 13, 15)
      .getValues();
  }
  function getTotalSlots() {
    return SpreadsheetApp
      .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
      .getSheetByName('Form Responses 1')
      .getRange(4, 2, 4, 23)
      .getValues();
  }
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
  function makeClassName(i, j, dept, crew, nextCrew, a_c, b_d, order, data) {
    var className = "";
    var length;
    var column;
    switch(dept) {
      case 'RollFed':
        column = 7;
        break;
      case 'Inline':
        column = 2;
        break;
      case 'NextRollFed':
        column = 9;
        break;
      case 'NextInline':
        column = 4;
        break;
    }
    if (crew == 'A_C' || nextCrew == 'B_D') {
      if (a_c.indexOf(j) > -1) {
        if (order == 'first') {
          length = data[j][column];
        }
        if (order == 'second') {
          length = data[j+14][column];
        }
      }
      if (b_d.indexOf(j) > -1) {
        if (order == 'first') {
          length = data[j+7][column];
        }
        if (order == 'second') {
          length = data[j+21][column];
        }
      }
    }
    if (crew == 'B_D' || nextCrew == 'A_C') {
      if (a_c.indexOf(j) > -1) {
        if (order == 'first') {
          length = data[j+7][column];
        }
        if (order == 'second') {
          length = data[j+21][column];
        }
      }
      if (b_d.indexOf(j) > -1) {
        if (order == 'first') {
          length = data[j][column];
        }
        if (order == 'second') {
          length = data[j+14][column];
        }
      }
    }
    if (i >= length) {
      className = "outside";
    } else {
      className = "normal";
    }
    return className;
  }
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
  
  function filtered_person(dept, crew, date, data) {
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
        if (data[i][11] == '1' || (data[i][11] == 0 && date < new Date())) {
          if (!last || last > date) {
            var formated_date = Utilities.formatDate(date, "CST", "MM/dd/Y");
            if (start < date) {
              if (formated_absence.indexOf(formated_date) > -1) {
                person.push(data[i][1] + " " + data[i][2] + 'strike');  
              } else {
                person.push(data[i][1] + " " + data[i][2]);
              }
            } else if (start > date) {
              continue;
            } else {
              if (formated_absence && formated_absence.toString().indexOf(formated_date) > -1) {
                person.push(data[i][1] + " " + data[i][2] + 'ye_str');
              } else {
                person.push(data[i][1] + " " + data[i][2] + 'yellow')
              }
            }
          }
        }
      }
    }
    return person;
  }
  function make_persons(crew, dept, week, data, order) {
    var persons = [];
    if (crew == "A_C") {
      if (order == "first") {
        var crews = ["A Crew", "A Crew", "B Crew", "B Crew", "A Crew", "A Crew", "A Crew"];
      } else {
        var crews = ["C Crew", "C Crew", "D Crew", "D Crew", "C Crew", "C Crew", "C Crew"];
      }
    }
    if (crew == "B_D") {
      if (order == "first") {
        var crews = ["B Crew", "B Crew", "A Crew", "A Crew", "B Crew", "B Crew", "B Crew"];
      } else {
        var crews = ["D Crew", "D Crew", "C Crew", "C Crew", "D Crew", "D Crew", "D Crew"];
      }
    }
    for (var i = 0; i < 7; i++) {
      persons[i] = filtered_person(dept, crews[i], week[i], data);
    }
    return persons;
  }
  function get_first_length(firstSchedule) {
    var length = {};
    length.inCurAB = Math.max(firstSchedule[0][4], firstSchedule[1][2], firstSchedule[2][2], firstSchedule[3][2], firstSchedule[4][2], firstSchedule[5][2], firstSchedule[6][2], firstSchedule[7][2], firstSchedule[8][2], firstSchedule[9][2], firstSchedule[10][2], firstSchedule[11][2], firstSchedule[12][2], firstSchedule[13][2]);
    length.inCurCD = Math.max(firstSchedule[14][2], firstSchedule[15][2], firstSchedule[16][2], firstSchedule[17][2], firstSchedule[18][2], firstSchedule[19][2], firstSchedule[20][2], firstSchedule[21][2], firstSchedule[22][2], firstSchedule[23][2], firstSchedule[24][2], firstSchedule[25][2], firstSchedule[26][2], firstSchedule[27][2]);

    length.inNextAB = Math.max(firstSchedule[0][4], firstSchedule[1][4], firstSchedule[2][4], firstSchedule[3][4], firstSchedule[4][4], firstSchedule[5][4], firstSchedule[6][4], firstSchedule[7][4], firstSchedule[8][4], firstSchedule[9][4], firstSchedule[10][4], firstSchedule[11][4], firstSchedule[12][4], firstSchedule[13][4]);
    length.inNextCD = Math.max(firstSchedule[14][4], firstSchedule[15][4], firstSchedule[16][4], firstSchedule[17][4], firstSchedule[18][4], firstSchedule[19][4], firstSchedule[20][4], firstSchedule[21][4], firstSchedule[22][4], firstSchedule[23][4], firstSchedule[24][4], firstSchedule[25][4], firstSchedule[26][4], firstSchedule[27][4]);

    length.rollCurAB = Math.max(firstSchedule[0][7], firstSchedule[1][7], firstSchedule[2][7], firstSchedule[3][7], firstSchedule[4][7], firstSchedule[5][7], firstSchedule[6][7],firstSchedule[7][7], firstSchedule[8][7], firstSchedule[9][7], firstSchedule[10][7], firstSchedule[11][7], firstSchedule[12][7], firstSchedule[13][7]);
    length.rollCurCD = Math.max(firstSchedule[14][7], firstSchedule[15][7], firstSchedule[16][7], firstSchedule[17][7], firstSchedule[18][7], firstSchedule[19][7], firstSchedule[20][7], firstSchedule[21][7], firstSchedule[22][7], firstSchedule[23][7], firstSchedule[24][7], firstSchedule[25][7], firstSchedule[26][7], firstSchedule[27][7]);

    length.rollNextAB = Math.max(firstSchedule[0][9], firstSchedule[1][9], firstSchedule[2][9], firstSchedule[3][9], firstSchedule[4][9], firstSchedule[5][9], firstSchedule[6][9],firstSchedule[7][9], firstSchedule[8][9], firstSchedule[9][9], firstSchedule[10][9], firstSchedule[11][9], firstSchedule[12][9], firstSchedule[13][9]);
    length.rollNextCD = Math.max(firstSchedule[14][9], firstSchedule[15][9], firstSchedule[16][9], firstSchedule[17][9], firstSchedule[18][9], firstSchedule[19][9], firstSchedule[20][9], firstSchedule[21][9], firstSchedule[22][9], firstSchedule[23][9], firstSchedule[24][9], firstSchedule[25][9], firstSchedule[26][9], firstSchedule[27][9]);
    return length;
  }
  function create_table(persons, max_length) {
    var table = new Array(7);
    for (var i = 0; i < max_length; i++) {
      table[i] = new Array();
    }
    for (var i=0; i<max_length; i++) {
      for (var j=0; j<persons.length; j++) {
        if (persons[j][i]) {
          table[i][j] = persons[j][i];
        } else {
          table[i][j] = null;
        }
      }
    }
    return table;
  }
function makeArrayFromTable(dept, length_AB, length_CD) {
  var crew = detectCrew(new Date());
  var week = getWeek(1);
  var data = getDataSortName();
  
  var table_data = [];
  var week_head = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  var date_head = [];
  for (var i in week) {
    date_head.push(Utilities.formatDate(week[i], "CST", "MM/dd/YY"));
  }
  if (crew == "A_C") {
    var first_crew_head = ['A Crew', 'A Crew', 'B Crew', 'B Crew', 'A Crew', 'A Crew', 'A Crew'];
    var second_crew_head = ['C Crew', 'C Crew', 'D Crew', 'D Crew', 'C Crew', 'C Crew', 'C Crew'];
  }
  if (crew == "B_D") {
    var first_crew_head = ['B Crew', 'B Crew', 'A Crew', 'A Crew', 'B Crew', 'B Crew', 'B Crew'];
    var second_crew_head = ['C Crew', 'C Crew', 'D Crew', 'D Crew', 'C Crew', 'C Crew', 'C Crew'];
  }
  
  var top_persons = make_persons(crew, dept, week, data, 'first');
  var bottom_persons = make_persons(crew, dept, week, data, 'second');
  
  var top_table = create_table(top_persons, length_AB);
  var bottom_table = create_table(bottom_persons, length_CD);
  
  table_data.push(week_head, date_head, first_crew_head);
  for (var i in top_table) {
    table_data.push(top_table[i]);
  }
  table_data.push(second_crew_head)
  for (var i in bottom_table) {
    table_data.push(bottom_table[i])
  }
  return table_data;
}
function downloadData(){
  var spreadsheetId = '1Ojl9dNq24dm6JkgB5GEOVZlsTL7k5LNT2F1FIwDKuJ4';
  var sheetName = Utilities.formatDate(firstOfWeek(1), "CST", "MM/dd/YY")
  var activeSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var newSheet = activeSpreadsheet.getSheetByName(sheetName);
  
  var length = get_length(getTotalSlots());

  var rollFedArray = makeArrayFromTable("Roll Fed", length.rollFedCurAB, length.rollFedCurCD);
  var inlineArray = makeArrayFromTable("Inline", length.inlineCurAB, length.inlineCurCD);
  var ecoStarArray = makeArrayFromTable("Eco Star", length.ecoStarCurAB, length.ecoStarCurCD);
  var northPlantArray = makeArrayFromTable("North Plant", length.northPlantCurAB, length.northPlantCurCD);

  var rollFedLen = length.rollFedCurAB + length.rollFedCurCD + 4;
  var inlineLen = length.inlineCurAB + length.inlineCurCD + 4;
  var ecoStarLen = length.ecoStarCurAB + length.ecoStarCurCD + 4;
  var northPlantLen = length.northPlantCurAB + length.northPlantCurCD + 4;
  
  var rollFedRange = 'B2:H' + (rollFedLen + 1);
  var inlineRange = 'B'+ (rollFedLen + 2) + ':H' + (rollFedLen + inlineLen + 1);
  var ecoStarRange = 'B'+ (rollFedLen + inlineLen + 3) + ':H' + (rollFedLen + inlineLen + ecoStarLen + 2);
  var northPlantRange = 'B'+ (rollFedLen + inlineLen + ecoStarLen + 4) + ':H' + (rollFedLen + inlineLen + ecoStarLen + northPlantLen + 3);

  if (newSheet == null) {
    newSheet = activeSpreadsheet.insertSheet();
    newSheet.setName(sheetName);
    //RollFed
    newSheet.getRange('A1').setValue('RollFed');
    newSheet.getRange(rollFedRange).setValues(rollFedArray);
    //Inline
    newSheet.getRange('A'+ (rollFedLen + 1)).setValue('Inline');
    newSheet.getRange(inlineRange).setValues(inlineArray);
    //Eco Star
    newSheet.getRange('A'+ (rollFedLen + inlineLen + 2)).setValue('Eco Star');
    newSheet.getRange(ecoStarRange).setValues(ecoStarArray);
    //North Plant
    newSheet.getRange('A'+ (rollFedLen + inlineLen + ecoStarLen + 3)).setValue('North Plant');
    newSheet.getRange(northPlantRange).setValues(northPlantArray);
    return true;
  } else {
    return false;
  }
}