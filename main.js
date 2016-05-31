function handleFile(e) {
 var files = e.target.files;
 var i,f;
 for (i = 0, f = files[i]; i != files.length; ++i) {
   var reader = new FileReader();
   var name = f.name;
   reader.onload = function(e) {
		 var data = e.target.result;
     workbook = XLSX.read(data, {type: 'binary'});
     mySheets = workbook.SheetNames;
	 };
 };
 reader.readAsBinaryString(f);
}

function handleMonth(e){
 if (workbook){
   currentMonth = e.target.value;
   currentMonthSheet = workbook.Sheets[currentMonth];
   monthNum = mySheets.indexOf(currentMonth) + 1;
   numDays = daysInMonth(monthNum, year);
   h = {};
   for (var j = 1; j <= numDays; j++){
     for (var key in currentMonthSheet) {
       if (currentMonthSheet[key].v == j) {
         if (h[j] && (j > numDays - 7)) {
            h[j] = key;
         }
         else if (!h[j]){
            h[j] = key;
         }
       }
     }
   }
 }
}
var yourFile = document.getElementById("your-files");
sel.addEventListener('change', handleMonth, false);
yourFile.addEventListener('change', handleFile, false);

function handleDrop(e) {
  e.stopPropagation();
  e.preventDefault();
  var files = e.dataTransfer.files;
  var i,f;
  for (i = 0, f = files[i]; i != files.length; ++i) {
    var reader = new FileReader();
    var name = f.name;
    reader.onload = function(e) {
      var data = e.target.result;

      /* if binary string, read with type 'binary' */
      var workbook = XLSX.read(data, {type: 'binary'});

      /* DO SOMETHING WITH workbook HERE */
    };
    reader.readAsBinaryString(f);
  }
}
var drog = document.getElementById("drog");
drog.addEventListener('drop', handleDrop, false);

// the function is going to use regular expression to compare two names
// name2 is input name, name1 is the name in excel file
// function{Boolean}
var nameCompare = function (name1, name2) {
  var nameOne = name1.toLowerCase().split(" ");
  var nameTwo = name2.toLowerCase().split(" ");
  if (nameOne[1] && !nameOne[1].includes("(") && nameOne[1].includes(".")) {
    nameOne[1] = nameOne[1].replace(".", "");
  }
  if (nameOne[0] === nameTwo[0]) {
    if (!nameOne[1]) {
      return true;
    }
    else if (!nameOne[1].includes("(") && !nameTwo[1].includes(nameOne[1])) {
      return false;
    }
    else {
      return true;
    }
  }
  else {
    return false;
  }
}

// first= First Name, last= Last Name, obj=workbook, schedule=Result Object
// the function is modifying the schedule Object
// the function doesn't return anything
var searchScheduleByName = function(first, last, startColIdx, endColIdx, startRow, endRow, date, schedule) {
  var COL = [A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB];
  var row, col, idx, lastIdx, newIdx;
  var name = first+" "+last;
  while (startRow < endRow ) {
    while (startColIdx < COL.length) {
      row = startRow.toString();
      col = COL[startColIdx]

      if (nameCompare(name, obj[idx])) {
        lastIdx = COL[startColIdx-1]+row;

        startColIdx ++;

        col = COL[startColIdx];
        newIdx = col+row;
        schedule[date] = [obj[lastIdx], obj[newIdx]];
      }
    }
  }
}
