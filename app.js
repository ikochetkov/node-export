const port = '3001'

const stylers = require('./styleObjects.js')
const xl = require('excel4node')
const express = require('express')
const cors = require('cors')


const app = express()
app.use(cors())
app.use(express.json()) // for parsing application/json
app.use(express.urlencoded({ extended: true })) // for parsing application/x-www-form-urlencoded

app.post('/palmsnow/statements', function (req, res, next) {

  const wb = new xl.Workbook();
  const ws = wb.addWorksheet('Statements');
  const { body: { exportRows, exportFields, excelName } } = req
  const headerStyleObj = wb.createStyle(stylers.header)
  const regularStyleObj = wb.createStyle(stylers.regular)
  const regularHighlightedStyleObj = wb.createStyle(stylers.regular_highlighted)

  exportFields.forEach((col, index) => {
    ws.cell(1, index + 1).style(headerStyleObj);
    ws.cell(1, index + 1).string(col.label);
    ws.column(index + 1).setWidth(index ? 20 : 30)
  })

  exportRows.forEach((row, index) => {
    const rowNumber = index + 2;
    const styling = row.Description.startsWith('*') ? regularHighlightedStyleObj : regularStyleObj
    Object.keys(row).forEach((key, keyIndex) => {
      ws.cell(rowNumber, keyIndex + 1).string(row[key])
      ws.cell(rowNumber, keyIndex + 1).style(styling)
      if (keyIndex == 0) ws.cell(rowNumber, keyIndex + 1).style({
        alignment: {
          horizontal: 'left',
          vertical: 'center'
        }
      })
    })
  })

  wb.write(`${excelName}.xlsx`, res);
  console.log(`File generated - ${new Date()}`)
})

app.listen(port, function () {
  console.log(`App is currently running at ${port} port!`);
});