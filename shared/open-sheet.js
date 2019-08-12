import { glueReadyPromise } from "./initialize-glue.js";
import { cloneObject } from "./utils.js";

let excelOptions = {
  clearGrid: true,
  worksheet: 'Data',
  inhibitLocalSave: false,
  // response: 'image',
  // topLeft: 'A1',
  // updateTrigger: { row: true, save: true},
  workbook: 'Portfolio Info',
  disableErrorViewer: false,
  window: 'top'
}

const columnConfig = [
  {fieldName: 'ric', header: 'Instrument', width: 10},
  {fieldName: 'description', header: 'Description', width: 20},
  {fieldName: 'price', header: 'Price'},
  {fieldName: 'shares', header: 'Number of shares', width: 16}
]

function openSheet(contact, callback) {
  return glueReadyPromise.then(glue => {
    let sheetConfig = {
      options: excelOptions,
      columnConfig,
      data: cloneObject(contact.context.portfolio)
    }

    sheetConfig.options.workbook = `${contact.displayName} Portfolio`

    return glue.excel.openSheet(sheetConfig)
      .then(sheet => {
        return sheet.onChanged((data, errorCb, successCb) => {
          let errors = checkSheetChangeForErrors(data);
          if (errors.length === 0) {
            console.log('successCb');
            successCb();
            callback(data);
          } else {
            console.log('errorCb');
            errorCb(errors);
          }
        })
      })
  })
}

function checkSheetChangeForErrors(data) {
  let errors = [];

  data.forEach((row, index) => {
    if (typeof row.price !== 'number' || row.price < 1) {
      errors.push({row: index + 1, column: 3, foregroundColor: 'white', backgroundColor: 'red', description: `"${row.price}" is not a valid positive number`})
    }

    if (typeof row.shares !== 'number' || row.shares < 1) {
      errors.push({row: index + 1, column: 4, foregroundColor: 'white', backgroundColor: 'red', description: `"${row.shares}" is not a valid positive number`})
    }
  })

  return errors;
}


export { openSheet }