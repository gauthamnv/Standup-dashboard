// Initialize wb as a global variable
var wb;
var standStatus = document.getElementById("standStatus").value;

function getByGroups() {
  const inputElements = document.querySelectorAll("form input");
  let groupedElements = new Array();
  inputElements.forEach((input) => {
    console.log(input);
    const baseId = input.getAttribute("id");
    if (baseId) {
      let splitId = baseId.split("_");
      let userName = splitId[0];
      let fieldType = splitId[1];

      let formValue = input.value;
      if (typeof groupedElements[userName] === "undefined") {
        groupedElements[userName] = new Array();
      }

      switch (fieldType) {
        case "name":
          groupedElements[userName]["name"] = formValue;
          break;
        case "project1":
          groupedElements[userName]["project1"] = formValue;
          break;
        case "project2":
          groupedElements[userName]["project2"] = formValue;
          break;
        case "project3":
          groupedElements[userName]["project3"] = formValue;
          break;
        case "project4":
          groupedElements[userName]["project4"] = formValue;
          break;
        case "project5":
          groupedElements[userName]["project5"] = formValue;
          break;
        case "project6":
          groupedElements[userName]["project6"] = formValue;
          break;
        case "project7":
          groupedElements[userName]["project7"] = formValue;
          break;
      }
    }
  });
  return groupedElements;
}
// var formData;
function processForm() {
  let formData = getByGroups();

  let capturedData = capturedata();

  capturedData.forEach((row) => {
    let name = row[0];
    if (!formData[name]) {
      formData[name] = {};
      formData[name]["name"] = row[0];
      formData[name]["project1"] = row[1];
      formData[name]["project2"] = row[2];
      formData[name]["project3"] = row[3];
      formData[name]["project4"] = row[4];
      formData[name]["project5"] = row[5];
      formData[name]["project6"] = row[6];
      formData[name]["project7"] = row[7];
    } else {
      let isDataRepeated = Object.values(formData[name]).every(
        (value, i) => value === row[i]
      );
      if (!isDataRepeated) {
        formData[name]["name"] = row[0];
        formData[name]["project1"] = row[1];
        formData[name]["project2"] = row[2];
        formData[name]["project3"] = row[3];
        formData[name]["project4"] = row[4];
        formData[name]["project5"] = row[5];
        formData[name]["project6"] = row[6];
        formData[name]["project7"] = row[7];
      }
    }
  });

  let dbData = new Array();
  dbData[0] = [
    "Name",
    "Project1",
    "Project2",
    "Project3",
    "Project4",
    "Project5",
    "Project6",
    "Project7",
  ];
  let count = 1;
  for (var key in formData) {
    dbData[count] = [];
    dbData[count][0] = formData[key]["name"];
    dbData[count][1] = formData[key]["project1"];
    dbData[count][2] = formData[key]["project2"];
    dbData[count][3] = formData[key]["project3"];
    dbData[count][4] = formData[key]["project4"];
    dbData[count][5] = formData[key]["project5"];
    dbData[count][6] = formData[key]["project6"];
    dbData[count][7] = formData[key]["project7"];
    count++;
  }
  exportToExcel(dbData);
}

const myButton = document.getElementById("addStatus");
myButton.addEventListener("click", processForm);
var currentDate = getDate();
document.getElementById("todayDate").innerHTML = currentDate;

// Check if the Excel file exists
function checkFileExists() {
  var fileExists = false;

  if (wb) {
    // Check if the workbook has a sheet named 'Sheet1'
    var sheetNames = wb.SheetNames;
    fileExists = sheetNames.includes("Status");
  }

  return fileExists;
}

function exportToExcel(formData) {
  console.log(formData);
  var ws = XLSX.utils.aoa_to_sheet(formData);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Status1");
  var todaysDate = getDate();
  if (standStatus == "standup") {
    var fileName = "projectStandup_" + todaysDate + ".xlsx";
  } else {
    var fileName = "projectStatus_" + todaysDate + ".xlsx";
  }

  XLSX.writeFile(wb, fileName);
}

function getDate() {
  var currentDate = new Date();

  var day = currentDate.getDate();
  var month = currentDate.getMonth() + 1; // Months are zero-based, so we add 1
  var year = currentDate.getFullYear();

  day = day < 10 ? "0" + day : day;
  month = month < 10 ? "0" + month : month;

  var formattedDate = day + "-" + month + "-" + year;
  return formattedDate;
}

// adding a row to the table
window.addRow = function () {
  var table = document
    .getElementById("todayProjectStatus")
    .getElementsByTagName("tbody")[0];
  var newRow = table.insertRow(table.rows.length);
  var cell1 = newRow.insertCell(0);
  var cell2 = newRow.insertCell(1);
  var cell3 = newRow.insertCell(2);
  var cell4 = newRow.insertCell(3);
  var cell5 = newRow.insertCell(4);
  var cell6 = newRow.insertCell(5);
  var cell7 = newRow.insertCell(6);
  var cell8 = newRow.insertCell(7);
  cell1.innerHTML =
    '<input type="text" name="name[]" placeholder="Enter Name">';
  cell2.innerHTML =
    '<input type="text" name="project1[]" placeholder="Enter Project 1">';
  cell3.innerHTML =
    '<input type="text" name="project2[]" placeholder="Enter Project 2">';
  cell4.innerHTML =
    '<input type="text" name="project3[]" placeholder="Enter Project 3">';
  cell5.innerHTML =
    '<input type="text" name="project4[]" placeholder="Enter Project 4">';
  cell6.innerHTML =
    '<input type="text" name="project5[]" placeholder="Enter Project 5">';
  cell7.innerHTML =
    '<input type="text" name="project6[]" placeholder="Enter Project 6">';
  cell8.innerHTML =
    '<input type="text" name="project7[]" placeholder="Enter Project 7">';
};

// capturing the data from the table and add to the existing formdata
function capturedata() {
  var table = document
    .getElementById("todayProjectStatus")
    .getElementsByTagName("tbody")[0];
  var rows = table.rows;
  var data = [];
  for (var i = 0; i < rows.length; i++) {
    data[i] = [];
    for (var j = 0; j < rows[i].cells.length; j++) {
      data[i][j] = rows[i].cells[j].getElementsByTagName("input")[0].value;
    }
  }
  return data;
}
