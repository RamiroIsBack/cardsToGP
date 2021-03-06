////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////// FUNCTIONS FOR MOTHER /////////////////////////////////////////////////
module.exports.getPath = () => {
  let path = "../";
  return path;
};

module.exports.groupItemNumbersByFormula = worksheetRead => {
  let formulaCol = worksheetRead.getColumn("A");
  let identificationCol = worksheetRead.getColumn("B");
  let itemNumberCol = worksheetRead.getColumn("C");
  let formulaSelected = formulaCol.values[2].result
    ? formulaCol.values[2].result
    : formulaCol.values[2];
  let itemsArray = [];
  let arrayOfGroupedObjects = [];

  for (let i = 2; i < formulaCol.values.length; i++) {
    if (formulaCol.values[i]) {
      let formulaColDependingOnIndex = formulaCol.values[i].result
        ? formulaCol.values[i].result
        : formulaCol.values[i];
      if (formulaColDependingOnIndex !== formulaSelected) {
        //meter este grup en el array de grupos
        arrayOfGroupedObjects.push({
          formula: formulaSelected,
          itemsArray
        });
        //start with the next formula
        formulaSelected = formulaColDependingOnIndex;
        itemsArray = [];
      }
      itemsArray.push({
        itemNumber: itemNumberCol.values[i].result
          ? itemNumberCol.values[i].result
          : itemNumberCol.values[i],
        identification: identificationCol.values[i].result
          ? identificationCol.values[i].result
          : identificationCol.values[i],
        formula: formulaSelected
      });
    }
  }
  return arrayOfGroupedObjects;
};

module.exports.writeDataInMother = (results, worksheetWrite) => {
  // E stokquantity F lot G exp H bin
  worksheetWrite.columns = [
    { header: "Formula", key: "formula", width: 20 },
    { header: "Item Number", key: "itemNumber", width: 20 },
    { header: "Stock Quantity", key: "stockQuantity", width: 20 },
    { header: "Lot", key: "lot", width: 20 },
    { header: "Exp", key: "exp", width: 20 },
    { header: "Bin", key: "bin", width: 20 },
    { header: "Type for Bin", key: "typeForBin", width: 20 }
  ];

  results.map((resultado, index) => {
    if (resultado.length > 0) {
      //hay datos q meter
      rowNumberToWrite = 0;
      resultado.forEach(element => {
        // Add a couple of Rows by key-value, after the last current row, using the column keys

        if (element.error) {
          //ponemos el error
          let rowToWrite = worksheetWrite.addRow({ itemNumber: element.error });
          rowToWrite.commit();
        } else {
          //escribimos los datos
          let rowToWrite = worksheetWrite.addRow({
            formula: element.formula,
            itemNumber: element.itemNumber,
            stockQuantity: element.stockQuantity,
            lot: element.lot,
            exp: element.expirationDate,
            bin: element.binLocation,
            typeForBin: element.typeForBin
          });
          rowToWrite.commit();
        }
      });
    } else {
      //no habia datos diferentes a 0
    }
  });
};

////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////// FUNCTIONS FOR CHILD //////////////////////////////////////////////////

module.exports.getExpeditionsColumn = rowToFindExp => {
  var found = 0;
  for (let i = 2; i < 100; i++) {
    val = rowToFindExp.values[i];
    if (val) {
      if (val.result) {
        var expPart = val.result.toString();
        expPart = expPart.substr(0, 3).toLowerCase();
        if (expPart === "exp" || expPart === "inv") {
          found = i;
          break;
        }
      }
    }
  }
  return found;
};

module.exports.getBinLocation = (binWorkSheet, typeForBin) => {
  var found = "";
  rowFindTypeForBin = binWorkSheet.getRow(5);
  rowForBinLocation = binWorkSheet.getRow(4);
  rowForBinLocationSecundary = binWorkSheet.getRow(3);
  for (let i=1 ;i< rowFindTypeForBin.values.length ; i++){
    let val = rowFindTypeForBin.values[i];
    if (val && val !== undefined) {
      let noSpacesOrLineBreaksVal = val
        .toString()
        .toLowerCase()
        .replace(/\s/g, "");
      let noSpacesOrLineBreaksTypeForBin = typeForBin
        .toString()
        .toLowerCase()
        .replace(/\s/g, "");
      if (noSpacesOrLineBreaksTypeForBin === noSpacesOrLineBreaksVal) {
        found = rowForBinLocation.values[i];
        if (found === "" || found === undefined) {
          found = rowForBinLocationSecundary.values[i];
          if (found === "" || found === undefined) {
            found = "not found";
          }
        }
        break;
      }
    }
  }
  // rowFindTypeForBin.values.forEach((val, i) => {
  // });
  if (found === "") {
    found =
      "There is no match between inventory Master Type for bin and any type in line 5 of the corresponding sheet";
  }
  return found;
};
module.exports.doNothing = ()=>{
  return 1;
}
