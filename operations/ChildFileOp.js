var Excel = require("exceljs");
var fs = require("fs");

var helpingFunctions = require("./HelpingFunctions");

module.exports.getDataFromFileChild = async document => {
  //open child document based on formula
  var dirPath = helpingFunctions.getPath(); //directory path
  var file = `${dirPath}${document.trim()}`;


  var workbookChild = new Excel.Workbook();
  //return workbookChild;
  try {
    await workbookChild.xlsx
      .readFile(file);
    dataToSendBack = getData(workbookChild, file);
    return dataToSendBack;
  } catch (error_1) {
    if (file === "") {
      dataToSendBack = [{
        error: ` there is no file for this formula to get the data from: ${formulaPartOfName}`
      }];
      return dataToSendBack;
    } else {
      dataToSendBack = [{
        error: `error: ${error_1.message} `
      }];
      return dataToSendBack;
    }
  }
};

const getData = (workbookChild, file) => {
  var inventoryWorksheet = workbookChild.getWorksheet("Inventory Master");

  if (!inventoryWorksheet) {
    let errorLocation = file;
    throw new Error(
      `${
        errorLocation
      } there is no Inventory Master sheet`
    );
  }

  var arrayOfDataFromChildObjects = []; // will contain all data from this child
  var row = inventoryWorksheet.getRow(3);
  if (row.values.length === 0) {
    let errorLocation = file;
    throw new Error(
      `${
        errorLocation
      } there is nothing on line 3 of inventory master`
    );
  }
  var itemNumber = null;
  var formula = file.substr(3, file.indexOf(" "));;
  var rowIndex = 0;

  for (rowIndex = 0; rowIndex < row.values.length; rowIndex++) {
    var itemNum = row.values[rowIndex];
    if (itemNum && itemNum !== undefined) {
      if (itemNum.toString().trim() === "?") {

        itemNumber = "?";

        fillOutChildData(arrayOfDataFromChildObjects, inventoryWorksheet, itemNumber, formula, workbookChild, rowIndex);
      } else {

        itemNumber = itemNum.toString().trim();

        fillOutChildData(arrayOfDataFromChildObjects, inventoryWorksheet, itemNumber, formula, workbookChild, rowIndex);


      }
    }
  }
  let r = 0
  return arrayOfDataFromChildObjects;
};

const fillOutChildData = (arrayOfDataFromChildObjects, inventoryWorksheet, itemNumber, formula, workbookChild, rowIndex) => {
  var unitNumberInStock = null;
  var expirationDate = null;
  var binLocation = null;
  var typeForBin = null;
  if (itemNumber !== null || itemNumber === "?") {
    // there is a coincidence in mother - child
    var matchingElementCol = inventoryWorksheet.getColumn(rowIndex); //getting col for quantity
    var lotCol = inventoryWorksheet.getColumn("A"); //getting col for lot
    typeForBin = matchingElementCol.values[4];
    var lotNotTotals = "";
    for (let i = 5; i < matchingElementCol.values.length; i++) {
      if (matchingElementCol.values[i]) {
        if (lotCol.values[i] && lotCol.values[i] !== undefined) {
          lotNotTotals = lotCol.values[i].result ?
            lotCol.values[i].result.toString() :
            lotCol.values[i].toString();
          //exclude last row with total amount
          if (lotNotTotals.toLowerCase() === "totals") {
            break; // there is no more usefull data
          }
        }

        unitNumberInStock = matchingElementCol.values[i].result !== undefined ? matchingElementCol.values[i].result : matchingElementCol.values[i];
        if (unitNumberInStock !== 0 && unitNumberInStock !== undefined) {
          //this is the line to get the data from

          var rowToFindExp = inventoryWorksheet.getRow(i);
          var expColNumber = helpingFunctions.getExpeditionsColumn(
            rowToFindExp
          );

          binLocation = "";
          //lotNotTotals contained in the worksheet name => open worksheet not force error
          let sheetNameForBin = "";
          for (let i = 0; i < workbookChild._worksheets.length; i++) {
            if (workbookChild._worksheets[i]) {
              let sheet = workbookChild._worksheets[i];

              if (
                sheet.name.trim() === lotNotTotals.trim()
                // this was catching more than expected        || sheet.name.includes(lotNotTotals.trim())
              ) {
                sheetNameForBin = sheet.name;
                break;
              }
            }
          }
          if (sheetNameForBin === "") {
            for (let i = 0; i < workbookChild._worksheets.length; i++) {
              if (workbookChild._worksheets[i]) {
                let sheet = workbookChild._worksheets[i];
                let firstPartOfLotNotTotals = lotNotTotals
                  .trim()
                  .substring(0, lotNotTotals.indexOf(" "));
                if (
                  firstPartOfLotNotTotals.length > 4 &&
                  sheet.name.includes(firstPartOfLotNotTotals)
                ) {
                  sheetNameForBin = sheet.name;
                  break;
                }
              }
            }
          }
          let binWorkSheet = workbookChild.getWorksheet(sheetNameForBin);
          if (!binWorkSheet) {
            binLocation = `${lotNotTotals} there is no corresponding worksheet for this lotNumber`;
          } else {
            binLocation = helpingFunctions.getBinLocation(
              binWorkSheet,
              typeForBin
            );
          }

          expirationDate = "";
          if (expColNumber === 0) {
            expirationDate = "expedition-date column not found";
          } else {
            var expCol = inventoryWorksheet.getColumn(expColNumber);
            if (expCol.values[i]) {
              expirationDate = expCol.values[i].result ?
                expCol.values[i].result :
                expCol.values[i];
            } else {
              expirationDate = "expiration-date doesnt exist for this one";
            }
          }

          arrayOfDataFromChildObjects.push({
            formula,
            itemNumber,
            stockQuantity: unitNumberInStock,
            lot: lotNotTotals,
            expirationDate,
            binLocation,
            typeForBin
          });

        }
      }
    }
    return;

  }
};