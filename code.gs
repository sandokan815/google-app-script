  function doGet() {
    return HtmlService.createTemplateFromFile('index.html')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  function getEmployeesByDate() {
    return SpreadsheetApp
      .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
      .getSheetByName('Form Responses 1')
      .getRange('A11:K')
      .sort(7)
      .getValues();
  }
  function getEmployeesByName() {
    return SpreadsheetApp
      .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
      .getSheetByName('Form Responses 1')
      .getRange('A11:K')
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
  function getMonday(week) {
    if (week == 'current') {
      var diff = 1;
    }
    if (week == 'next') {
      var diff = 8;
    }
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
  function detectCrew(date) {
    var first = new Date(date.getFullYear(), 0, 1);
    var weekNo = Math.ceil( (((date - first) / 86400000) + first.getDay() + 1) / 7 );
    if (date.getDay() == 0) {
      weekNo = weekNo - 1;
    }
    if (weekNo % 2 == 0) {
      return "A_C";
    } else {
      return "B_D";
    }
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
      if (start && data[i][3] == dept && data[i][4] == crew){
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