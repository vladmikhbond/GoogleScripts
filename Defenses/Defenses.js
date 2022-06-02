/** @OnlyCurrentDoc */
/* Макрос для визначення дат захисту

Дані розташовані на активному листі у двох регіонах, поділених щонайменше двома пустими стовбцями.
У першому регіоні у першому рядку обов'язково мають бути такі назви стовбців: Name, Rating, ProjectId, Desired.
У другому регіоні у першому рядку обов'язково мають бути такі назви стовбців: Date, Capacity.
У першому регіоні порядок рядків значення не має, у другому рядки повинні бути впорядковані по зростанню дат.
Форматування не має значення в обох регіонах.

Алгоритм
1 Корегуємо об'єм для членів команд (додати окремий стовпчик) 
  Команда представляється капітаном (у кого найвищий р.) з об'ємом, який дорівнює розміру команди.
  Об'єм інших членів команди = 0. Студенти з нульовим обємом з розподіла тимчасово виключаються.
  Якщо капітан не заявив бажаної дати, вважаємо бажаною найранішу можливу дату.
2 Впорядковуємо студентів по зростанню рейтінга.
3 Встановлюємо реальну дату для кожного студента - найближчу до бажаної вільну дату.
4 Додаємо на заброньовані місця тимчасово виключених з розподілу членів команди.
5 Розкидуємо по вільних датах тих, хто не заявив бажаної дати захисту.
 */

function Defenses() 
{
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Indices in s (students) table
  let name = findColFor("Name"); 
  let raiting = findColFor("Rating"); 
  let projectId = findColFor("ProjectId"); 
  let desired = findColFor("Desired");
  let volume = desired + 1, real = desired + 2, id = desired + 3;
  // Indices in d (dates) table
  let date = 0;
  let capasity = 1;
  
  // Load values
  let s = sheet.getRange(1, name+1).getDataRegion().getValues();                 // s - students
  s.shift(); // remove headers
  s.forEach((x, i) => x.push(1, "", i))      // add 3 cols: volume, real, id

  // Load dates
  let d = sheet.getRange(1, findColFor("Date") + 1).getDataRegion().getValues(); // d - dates
  d.shift(); // remove headers

  // Calculate the lead's volumes
  sortV(raiting, "desc");
  sortV(projectId);
  for (let i = 0; i < s.length - 1; i++) {

    if (s[i][projectId] && s[i][projectId] == s[i+1][projectId]) {
      s[i+1][volume] += s[i][volume];
      s[i][volume] = 0;
    } ;
  }

  // Set real dates for good students and for teams
  sortV(raiting);
  for (let i = 0; i < s.length; i++) 
  {
    // Set real dates for team members only
    if (s[i][volume] == 0) {
      s[i][real] = s[i-1][real];
    }
    // Set real dates for good studs & lazy leads
    else
    {
      let desiredDate = s[i][desired];
      // Forse desired date for lazy leads
      if (!desiredDate && s[i][volume] > 1) {
        desiredDate = d[0][date]
      }
      // Calc start date index on desired date
      if (desiredDate) {
        let startDateIndex = d.findIndex(x => x[date].toString() == desiredDate.toString());
        if (startDateIndex == -1) 
          throw ("Wrong desired date for " + s[i][name]);
        s[i][real] = getDiffDate(s[i], startDateIndex );     
      }      
    }
  }

  // Set a real date for lazy students (earliest)
  for (let student of s)  {
    if (!student[real]) {
      let iDate = d.findIndex(x => x[capasity] > 0);
      if (iDate != -1) {
        d[iDate][capasity]--;
        student[real] = d[iDate][date];
      }
    }
  }

  // output s 
  sortV(id);
  sheet.getRange(1, desired + 2).setValue("Real").setFontWeight('bold');
  sheet.getRange(2, desired + 2, s.length, 1).setValues(s.map(x => [x[real]]));

  // output d
  ts = ["10:00","10:30","11:00","11:30","12:00","12:30","13:00","13:30","14:00","14:30","15:00","15:30","16:00","16:30"];
  sortV(real);
  let row = 2, col = desired + 7;
  sheet.getRange(1, col, 1, 2).setValues([["Date", "Name"]]).setFontWeight('bold');;
  for (const d1 of d) {
    // show date
    sheet.getRange(row++, col).setValue(d1[date]).setFontWeight('bold');
    let studs = s.filter(x => x[real] == d1[date]);
    t = 0;
    for (const student of studs) {
      // show time & student name
      sheet.getRange(row, col).setValue(ts[t++]);           
      sheet.getRange(row++, col+1).setValue(student[name]);
    }
  }
      
 
  // ------------------------ INNER FUNCTIONS -----------------------
  
 
  function sortV(i, desc) {
    if (desc)
      s.sort( (a, b) => b[i] - a[i])
    else
      s.sort( (a, b) => a[i] - b[i])
  }

  // Generate numbers: 0, 1,-1 ,2,-2, 3,-3,...   
  function delta(n) {
    return n % 2 ? (n + 1) / 2 : -n / 2;  
  }

  function getDiffDate(student, startDateIndex) 
  {
    for (let tryNo = 0; tryNo < 100; tryNo++) {
       let idx = startDateIndex + delta(tryNo); 
       if (idx < 0 || idx >= d.length || student[volume] > d[idx][capasity])
          continue;
       d[idx][capasity] -= student[volume];
       return d[idx][date];
    }
    throw ("Сannot find real date for " + student[name]);
  }

  function findColFor(sample) {
    const cs = sheet.getRange(1,1, 1, 20).getValues();
    let i = cs[0].indexOf(sample);
    if (i != -1) return i;
    throw ("Can't find sample " + sample);
  }

};
