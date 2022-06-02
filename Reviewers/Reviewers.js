/** @OnlyCurrentDoc */
/* 
Макрос для призначення рецензентів.
Додаткові дані:
  A1 - номер верхнього рядка,
  B1 - номер нижнього рядка,
  C1 - кількість стовпців,
  У стовбці H - позначки групових проектів 
  (члени однієї команди мають однакові позначки, члени різних команд - різні).
Критичним для програми є розташування стовбців: 1, 2 і 6.
Рядки даних впорядковані по стовбцю 1.

Натиснути Ctrl-Alt-Shift-1.
З'явиться новий стовбець із прізвищами рецензентів.
Для відкату - Ctrl-Z.
*/

function Reviewers() {
   // Indices in v table
   let _id = 0, _studName = 1, _prep = 5, _projId = 7,  /*  */ _projW = 8, _reviewer = 9;
 
   
   // Load values
   let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   let top = sheet.getRange("a1").getValue();
   let bottom = sheet.getRange("b1").getValue();
   let width = sheet.getRange("c1").getValue();
   let v = sheet.getRange(top, 1, bottom - top + 1, width).getValues(); 
 
   v.forEach((x, i) => x.push(0, ""))      // _projW = 8, _rever = 9;
 
   // set all projId
   sortV(_projId, 'desc');
   let i = 0;
   while (v[i][_projId]) i++;
   let k = v[0][_projId] + 1;
   for (; i < v.length; i++) {
     v[i][_projId] = k++;
   }
 
   // calc projest weights
   defineWeightV(_projId, _projW);
   // dictionary prep weights
   const prepWeights = {};
   for (let x of v) {
     let k = x[_prep];
     if (prepWeights[k]) prepWeights[k]++; else prepWeights[k] = 1;
   }
   // dictionary project weights
   const projectWeights = {};
   for (let x of v) {
     let k = x[_projId];
     if (projectWeights[k]) projectWeights[k]++; else projectWeights[k] = 1;
   }
   
   sortV(_projId);
   sortV(_projW, 'desc');
   shuffleV();
 
   for (let i1 = 0, i2 = 1; i1 < v.length; i1 = i2) 
   {
     // отделяем проекты   [i1, i2) - segment of project 
     i2 = i1 + 1; 
       
     while (v[i2] && v[i2][_projId] == v[i1][_projId]) 
       i2++;
     
     if (hasReviewer(i1, i2))
       continue;   // у проекта уже есть рецензент 
 
     // находим рецензента для всего проекта
     let idxs = findReviewer(i1, i2, v[i1][_projW]);
     if (idxs.length == 0) {
       console.log("i1, i2", i1, i2)
       continue;
     }
 
     // производим обмен рецензированием и уменьшаем веса обоих препов
     let idxPrep = v[idxs[0]][_prep];
     for (let i = i1, k = 0; i < i2; i++, k++) 
     {
       let iPrep = v[i][_prep];
       if (!v[idxs[k]][_reviewer] && !v[i][_reviewer]) {
         v[idxs[k]][_reviewer] = iPrep
         prepWeights[iPrep]--;
   
         v[i][_reviewer] = idxPrep;
         prepWeights[idxPrep]--;
       } 
       else 
         console.log ("place is occupated", v[idxs[k]][_reviewer], !v[i][_reviewer])  
     }
   }
 
   showResult()
 
   // ------------------------ INNER FUNCTIONS -----------------------
   
   function findReviewer(i1, i2, projW) 
   {
     if (i1 == 129)
       i1 = i1  
 
     for (let i = v.length - 1; i >= 0; i--) {
       let iPrep = v[i][_prep];
       if (prepWeights[iPrep] < projW) // не хватвает веса
          continue;
       if (v[i][_reviewer])            // уже женат       
          continue;
 
       let isCurator = false;          // один из рук. распределяемого проекта
       for (let j = i1; j < i2; j++) {
         if (iPrep == v[j][_prep]) {
           isCurator = true;
           break;
         }
       }
       // find some lines with
       if (isCurator) 
          continue;
       let result = [i];       
       let w = projW - 1;
       // находим след.экземпляры рецензента
       
       for( let j = i - 1; j >= 0 && w > 0; j--) {
         if (v[j][_prep] != iPrep || v[j][_reviewer]) 
           continue;
         result.push(j);
         w--;
       } 
       if (w == 0 )
         return result;       
     }
     return [];
   }
 
   function hasReviewer(i1, i2) {
       for (let i = i1; i < i2; i++) {
         if (!v[i][_reviewer]) return false;
       }
       return true;
   }
 
   // shuffle simple projects 
   function shuffleV() { 
     let simpleStartIdx = v.findIndex(x => x[_projW] == 1);  
     let u = v.slice(simpleStartIdx);
     u.sort((a, b) => (a[_studName] < b[_studName] ? -1 : 1));
     v = v.slice(0, simpleStartIdx).concat(u);
   }
 
   function sortV(i, desc) {
     if (desc)
       v.sort( (a, b) => b[i] - a[i])
     else
       v.sort( (a, b) => a[i] - b[i])
   }
 
   function defineWeightV(id, w) {
   // sortV(id);
     v[0][w] = 1;
     for (let i = 1; i < v.length; i++) {
       if (v[i][id] == v[i - 1][id])
         v[i][w] = v[i - 1][w] + 1;
       else   
         v[i][w] = 1;
     }
     for (let i = v.length - 2; i >= 0;  i--) {
       if (v[i][id] == v[i + 1][id])
         v[i][w] = v[i + 1][w];
     }
     
   }
 
   function showResult()
   { 
     // for debug
     // var spreadsheet = SpreadsheetApp.getActive();
     // let newSheet = spreadsheet.insertSheet(1);
     // newSheet.getRange(top, 1, bottom - top + 1, v[0].length).setValues(v);
     // 
     sortV(_id);
     let a = v.map(x => [x[_reviewer]]);
     sheet.insertColumnAfter(width + 1);
     let style = sheet.getRange(_prep + 1, top + 1).getTextStyle();
     sheet.getRange(top, width + 1 , bottom - top + 1, 1).setValues(a).setTextStyle(style);
   }
 
 }
 