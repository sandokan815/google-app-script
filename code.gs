  function doGet() {
    return HtmlService.createTemplateFromFile('index.html')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
  }

  function getDataSortDate() {
    return SpreadsheetApp
      .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
      .getSheetByName('Form Responses 1')
      .getRange('A11:J')
      .sort(7)
      .getValues();
  }
  function getDataSortName() {
    return SpreadsheetApp
      .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
      .getSheetByName('Form Responses 1')
      .getRange('A11:J')
      .sort(3)
      .getValues();
  }

  function detectCrew(now) {
    var first = new Date(now.getFullYear(), 0, 1);
    var weekNo = Math.ceil( (((now - first) / 86400000) + first.getDay() + 1) / 7 );
    if (weekNo % 2 == 0) {
      return "A_C";
    } else {
      return "B_D";
    }
  }
  function firstOfWeek() {
    var now = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var monday = new Date(today.setDate(today.getDate()-today.getDay()+1));
    return monday;
  }
  function getWeek() {
    var curr = new Date();
    var first = curr.getDate() - curr.getDay() + 1;
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
      if (data[i][6]) {
        start = new Date(new Date(data[i][6]).getTime() + 2 * 60* 60* 1000);
      }
      if (data[i][9]) {
        last = new Date(new Date(data[i][9]).getTime() + 2 * 60* 60* 1000);
      }
      if (start && data[i][3] == dept && data[i][4] == crew){
        if (!last || last > date) {
          if (start < date) {
            person.push(data[i][1] + " " + data[i][2]);
          } else if (start > date) {
            continue;
          } else {
            person.push(data[i][1] + " " + data[i][2] + 'yellow');
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
  function get_max_length(persons) {
    max_length = 0;
    for (var i=0; i<persons.length; i++) {
      if (persons[i].length > max_length) {
        max_length = persons[i].length;
      }
    }
    return max_length;
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
