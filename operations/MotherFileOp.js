var Excel = require("exceljs");
var mapSeries = require("async/mapSeries");
var getDataFromFileChild = require("./ChildFileOp.js").getDataFromFileChild;
var helpingFunctions = require("./HelpingFunctions.js");

module.exports.readingMotherFile = documentList => {

  console.log("... start data processing ...");
  var dirPath = helpingFunctions.getPath(); //directory path
  //erase the folder containing the program and Stock Loading-Inter and FG.xlsx
  documentList = documentList.filter(
    doc =>
    doc !== "Stock Loading-Inter and FG.xlsx" && doc !== "results Cards to GP.xlsx" &&
    doc.substring(doc.length - 5) === ".xlsx"
  );

  //hacer un mapeado async de cada uno de los blokes
  mapSeries(
    documentList,
    (document, callback) => {

      getDataFromFileChild(document)
        .then(res => {
          callback(null, res);
        })
        .catch(err => {
          console.log(err);
          callback(err);
          throw err;
        });
    },
    function (err, results) {
      if (err) console.log(err); //TODO::: somenthing more??
      //results will be an array of objects
      writeData(results, dirPath);
    }
  );

};

var writeData = (results, dirPath) => {
  var workbookWrite = new Excel.Workbook();
  // create new sheet with pageSetup settings for A4 - landscape
  var worksheetWrite = workbookWrite.addWorksheet("sheet", {
    pageSetup: {
      paperSize: 9,
      orientation: "landscape"
    }
  });
  helpingFunctions.writeDataInMother(results, worksheetWrite);
  let resultFile = `./results Cards to GP.xlsx`;
  workbookWrite.xlsx.writeFile(resultFile).then(function () {
    console.log(`... done ...
      result file: ${resultFile}`);
  });
};