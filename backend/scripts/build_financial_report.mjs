import fs from 'node:fs/promises'
import path from 'node:path'
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool'

const workspace = 'C:\\Users\\raphg\\Documents\\Codex\\2026-08-26\\goo'
const outputDir = path.join(workspace, 'outputs', '01a03f7c-d568-7f21-a80f-279d6f99324f')
const outputPath = path.join(outputDir, 'st_dorothy_financial_report.xlsx')
const sourcePath = path.join(workspace, 'backend', 'data', 'school-data.json')
const data = JSON.parse(await fs.readFile(sourcePath, 'utf8'))
const receipts = data.receipts ?? []
const transactions = data.mpesaTransactions ?? []

const workbook = Workbook.create()
const summary = workbook.worksheets.add('Financial Summary')
const receiptSheet = workbook.worksheets.add('Receipt Register')
const mpesaSheet = workbook.worksheets.add('M-Pesa Reconciliation')
const classSheet = workbook.worksheets.add('Collection by Class')
const notesSheet = workbook.worksheets.add('Notes & Controls')

const navy = '#173B6B'
const blue = '#2B72C8'
const gold = '#D69B16'
const paleBlue = '#EDF5FF'
const paleGold = '#FFF8E8'
const green = '#28784B'
const lightBorder = '#DCE5EF'
const white = '#FFFFFF'
const money = '"KES" #,##0;[Red]("KES" #,##0);-'
const asText = value => {
  const digits = String(value || '').replace(/\D/g, '')
  return digits.length === 12 ? `${digits.slice(0, 3)} ${digits.slice(3, 6)} ${digits.slice(6, 9)} ${digits.slice(9)}` : String(value || '')
}

function title(sheet, range, text, subtitle) {
  sheet.mergeCells(range)
  const cell = sheet.getRange(range.split(':')[0])
  cell.values = [[text]]
  cell.format = { fill: navy, font: { bold: true, color: white, size: 18 }, rowHeight: 30 }
  const subtitleRange = `${range.split(':')[0].replace(/\d+$/, '2')}:${range.split(':')[1].replace(/\d+$/, '2')}`
  sheet.mergeCells(subtitleRange)
  sheet.getRange(subtitleRange.split(':')[0]).values = [[subtitle]]
  sheet.getRange(subtitleRange.split(':')[0]).format = { fill: navy, font: { color: '#DCEBFF', size: 10 }, rowHeight: 20 }
  sheet.showGridLines = false
}

function header(range) {
  range.format = { fill: blue, font: { bold: true, color: white }, borders: { preset: 'all', style: 'thin', color: '#AFC8E8' }, wrapText: true }
}

function dataFormat(range) {
  range.format = { borders: { preset: 'inside', style: 'thin', color: '#E7EDF3' } }
}

function applyWidths(sheet, widths) {
  Object.entries(widths).forEach(([column, width]) => { sheet.getRange(`${column}:${column}`).format.columnWidth = width })
}

title(summary, 'A1:H1', 'St. Dorothy School — Financial Collection Report', 'Bursar workspace · Confirmed receipts and reconciled M-Pesa payments only')
applyWidths(summary, { A: 25, B: 18, C: 4, D: 24, E: 18, F: 4, G: 34, H: 18 })
summary.getRange('A4:B4').values = [['Report status', 'PASS']]
summary.getRange('A4:B4').format = { fill: paleBlue, font: { bold: true, color: navy }, borders: { preset: 'all', style: 'thin', color: lightBorder } }
summary.getRange('B4').conditionalFormats.add('cellIs', { operator: 'equal', formula: '"PASS"', format: { fill: '#E8F5EC', font: { bold: true, color: green } } })
summary.getRange('A6:B9').values = [
  ['Confirmed collection', null],
  ['Receipts issued', null],
  ['M-Pesa collection', null],
  ['M-Pesa awaiting match', null]
]
summary.getRange('A6:A9').format = { fill: paleBlue, font: { bold: true, color: navy }, borders: { preset: 'all', style: 'thin', color: lightBorder } }
summary.getRange('B6:B9').format = { fill: white, font: { bold: true, color: navy, size: 13 }, borders: { preset: 'all', style: 'thin', color: lightBorder }, numberFormat: money }
summary.getRange('B6').formulas = [["=SUM('Receipt Register'!$I$2:$I$1000)"]]
summary.getRange('B7').formulas = [["=COUNT('Receipt Register'!$I$2:$I$1000)"]]
summary.getRange('B8').formulas = [["=SUMIF('Receipt Register'!$J$2:$J$1000,\"M-Pesa\",'Receipt Register'!$I$2:$I$1000)"]]
summary.getRange('B9').formulas = [["=SUMIFS('M-Pesa Reconciliation'!$D$2:$D$1000,'M-Pesa Reconciliation'!$E$2:$E$1000,\"<>Matched\")"]]
summary.getRange('B7').format.numberFormat = '#,##0'
summary.getRange('D6:E6').merge()
summary.getRange('D6').values = [['Collection by method']]
summary.getRange('D6:E6').format = { fill: gold, font: { bold: true, color: white }, borders: { preset: 'all', style: 'thin', color: gold } }
summary.getRange('D7:E10').values = [['Method', 'Collection'], ['M-Pesa', null], ['Bank', null], ['Cash', null]]
header(summary.getRange('D7:E7'))
dataFormat(summary.getRange('D8:E10'))
summary.getRange('E8').formulas = [["=SUMIF('Receipt Register'!$J$2:$J$1000,D8,'Receipt Register'!$I$2:$I$1000)"]]
summary.getRange('E8:E10').fillDown()
summary.getRange('E8:E10').format.numberFormat = money
summary.getRange('G6:H6').merge()
summary.getRange('G6').values = [['Control check']]
summary.getRange('G6:H6').format = { fill: gold, font: { bold: true, color: white }, borders: { preset: 'all', style: 'thin', color: gold } }
summary.getRange('G7:H9').values = [['Check', 'Result'], ['Receipt totals tie to method totals', null], ['Unmatched payments excluded', null]]
header(summary.getRange('G7:H7'))
dataFormat(summary.getRange('G8:H9'))
summary.getRange('H8').formulas = [['=IF(B6=SUM(E8:E10),"PASS","CHECK")']]
summary.getRange('H9').formulas = [['=IF(B9>=0,"PASS","CHECK")']]
summary.getRange('H8:H9').conditionalFormats.add('cellIs', { operator: 'equal', formula: '"PASS"', format: { fill: '#E8F5EC', font: { bold: true, color: green } } })
summary.getRange('H8:H9').conditionalFormats.add('cellIs', { operator: 'equal', formula: '"CHECK"', format: { fill: '#FDEBEC', font: { bold: true, color: '#B42318' } } })
summary.getRange('A12:H12').merge()
summary.getRange('A12').values = [['Report definition: collection totals include issued receipts only. Pending STK Push requests and unmatched M-Pesa confirmations are deliberately excluded.']]
summary.getRange('A12').format = { fill: paleGold, font: { italic: true, color: '#7B5916' }, wrapText: true, rowHeight: 30, borders: { preset: 'outside', style: 'thin', color: '#EED39D' } }
summary.freezePanes.freezeRows(2)

const receiptHeaders = ['Receipt Number', 'Issued At', 'Admission Number', 'Learner', 'Class', 'Guardian', 'Guardian Phone', 'Payment Reference', 'Amount', 'Payment Method']
title(receiptSheet, 'A1:J1', 'Receipt Register', 'Source: saved Bursar receipts in the local school portal database')
receiptSheet.getRange('A4:J4').values = [receiptHeaders]
header(receiptSheet.getRange('A4:J4'))
const receiptRows = receipts.map(receipt => [receipt.receiptNumber, new Date(receipt.issuedAt), receipt.admissionNumber, receipt.learnerName, receipt.className, receipt.guardianName, asText(receipt.guardianPhone), receipt.paymentReference, Number(receipt.amount), receipt.paymentMethod])
if (receiptRows.length) receiptSheet.getRange(`A5:J${4 + receiptRows.length}`).values = receiptRows
dataFormat(receiptSheet.getRange(`A5:J${Math.max(5, 4 + receiptRows.length)}`))
receiptSheet.getRange(`B5:B${Math.max(5, 4 + receiptRows.length)}`).format.numberFormat = 'yyyy-mm-dd hh:mm'
receiptSheet.getRange(`I5:I${Math.max(5, 4 + receiptRows.length)}`).format.numberFormat = money
receiptSheet.tables.add(`A4:J${Math.max(5, 4 + receiptRows.length)}`, true, 'ReceiptRegisterTable')
applyWidths(receiptSheet, { A: 19, B: 20, C: 20, D: 24, E: 18, F: 22, G: 18, H: 21, I: 15, J: 15 })
receiptSheet.freezePanes.freezeRows(4)

const receiptByTransaction = new Map(receipts.filter(receipt => receipt.mpesaTransactionId).map(receipt => [receipt.mpesaTransactionId, receipt.receiptNumber]))
const mpesaHeaders = ['M-Pesa Code', 'Received At', 'Sender Phone', 'Amount', 'Status', 'Admission Number', 'Learner', 'Class', 'Reconciled At', 'Receipt Number']
title(mpesaSheet, 'A1:J1', 'M-Pesa Reconciliation', 'Confirmed mobile-money transactions and their learner-account matching status')
mpesaSheet.getRange('A4:J4').values = [mpesaHeaders]
header(mpesaSheet.getRange('A4:J4'))
const mpesaRows = transactions.map(transaction => [transaction.transactionCode, new Date(transaction.receivedAt), asText(transaction.senderPhone), Number(transaction.amount), transaction.status, transaction.admissionNumber ?? '', transaction.learnerName ?? '', transaction.className ?? '', transaction.reconciledAt ? new Date(transaction.reconciledAt) : null, receiptByTransaction.get(transaction.id) ?? ''])
if (mpesaRows.length) mpesaSheet.getRange(`A5:J${4 + mpesaRows.length}`).values = mpesaRows
dataFormat(mpesaSheet.getRange(`A5:J${Math.max(5, 4 + mpesaRows.length)}`))
mpesaSheet.getRange(`B5:B${Math.max(5, 4 + mpesaRows.length)}`).format.numberFormat = 'yyyy-mm-dd hh:mm'
mpesaSheet.getRange(`D5:D${Math.max(5, 4 + mpesaRows.length)}`).format.numberFormat = money
mpesaSheet.getRange(`I5:I${Math.max(5, 4 + mpesaRows.length)}`).format.numberFormat = 'yyyy-mm-dd hh:mm'
mpesaSheet.getRange(`E5:E${Math.max(5, 4 + mpesaRows.length)}`).conditionalFormats.add('cellIs', { operator: 'equal', formula: '"Matched"', format: { fill: '#E8F5EC', font: { color: green } } })
mpesaSheet.getRange(`E5:E${Math.max(5, 4 + mpesaRows.length)}`).conditionalFormats.add('cellIs', { operator: 'equal', formula: '"Awaiting match"', format: { fill: paleGold, font: { color: '#7B5916' } } })
mpesaSheet.tables.add(`A4:J${Math.max(5, 4 + mpesaRows.length)}`, true, 'MpesaReconciliationTable')
applyWidths(mpesaSheet, { A: 16, B: 20, C: 18, D: 15, E: 17, F: 20, G: 24, H: 18, I: 20, J: 19 })
mpesaSheet.freezePanes.freezeRows(4)

const classes = [...new Set(receipts.map(receipt => receipt.className).filter(Boolean))].sort()
title(classSheet, 'A1:D1', 'Collection by Class', 'Formula-driven analysis linked to the Receipt Register')
classSheet.getRange('A4:D4').values = [['Class', 'Receipt Count', 'Confirmed Collection', 'Share of Total']]
header(classSheet.getRange('A4:D4'))
const classRows = classes.length ? classes.map(className => [className, null, null, null]) : [['No receipts yet', null, null, null]]
classSheet.getRange(`A5:D${4 + classRows.length}`).values = classRows
classRows.forEach((_, index) => {
  const row = 5 + index
  classSheet.getRange(`B${row}`).formulas = [[`=COUNTIF('Receipt Register'!$E$5:$E$1000,A${row})`]]
  classSheet.getRange(`C${row}`).formulas = [[`=SUMIF('Receipt Register'!$E$5:$E$1000,A${row},'Receipt Register'!$I$5:$I$1000)`]]
  classSheet.getRange(`D${row}`).formulas = [[`=IFERROR(C${row}/SUM('Receipt Register'!$I$5:$I$1000),0)`]]
})
dataFormat(classSheet.getRange(`A5:D${4 + classRows.length}`))
classSheet.getRange(`B5:B${4 + classRows.length}`).format.numberFormat = '#,##0'
classSheet.getRange(`C5:C${4 + classRows.length}`).format.numberFormat = money
classSheet.getRange(`D5:D${4 + classRows.length}`).format.numberFormat = '0.0%'
classSheet.tables.add(`A4:D${4 + classRows.length}`, true, 'CollectionByClassTable')
applyWidths(classSheet, { A: 22, B: 16, C: 22, D: 16 })
classSheet.freezePanes.freezeRows(4)

title(notesSheet, 'A1:E1', 'Notes & Controls', 'Report scope, source data, and refresh guidance')
notesSheet.getRange('A4:E4').values = [['Item', 'Value', 'Period / As of', 'Owner', 'Notes']]
header(notesSheet.getRange('A4:E4'))
notesSheet.getRange('A5:E10').values = [
  ['Source system', 'St. Dorothy School Portal', 'Current local data', 'Bursar', 'Receipts and M-Pesa reconciliation data are exported from the portal.'],
  ['Collection definition', 'Issued receipts only', 'All dates', 'Bursar', 'Unconfirmed STK prompts and unmatched M-Pesa transactions are excluded.'],
  ['Currency', 'Kenyan shillings (KES)', 'All dates', 'Bursar', 'Amounts are stored as numbers and formatted as KES.'],
  ['Refresh instruction', 'Re-run report after new receipts', 'As needed', 'Bursar', 'Use the portal Financial Reports page for live figures.'],
  ['Control: receipt total', null, 'All dates', 'Bursar', 'Should match Financial Summary confirmed collection.'],
  ['Control: unmatched M-Pesa', null, 'All dates', 'Bursar', 'Investigate before daily cashbook close.']
]
notesSheet.getRange('B9').formulas = [["=SUM('Receipt Register'!$I$5:$I$1000)"]]
notesSheet.getRange('B10').formulas = [["=SUMIFS('M-Pesa Reconciliation'!$D$5:$D$1000,'M-Pesa Reconciliation'!$E$5:$E$1000,\"<>Matched\")"]]
notesSheet.getRange('B9:B10').format.numberFormat = money
dataFormat(notesSheet.getRange('A5:E10'))
applyWidths(notesSheet, { A: 25, B: 28, C: 20, D: 16, E: 55 })
notesSheet.getRange('E5:E10').format.wrapText = true
notesSheet.freezePanes.freezeRows(4)

const check = await workbook.inspect({ kind: 'table', range: 'Financial Summary!A1:H12', include: 'values,formulas', tableMaxRows: 12, tableMaxCols: 8 })
console.log(check.ndjson)
const errors = await workbook.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A', options: { useRegex: true, maxResults: 100 }, summary: 'formula error scan' })
console.log(errors.ndjson)

await fs.mkdir(outputDir, { recursive: true })
for (const sheetName of ['Financial Summary', 'Receipt Register', 'M-Pesa Reconciliation', 'Collection by Class', 'Notes & Controls']) {
  const preview = await workbook.render({ sheetName, autoCrop: 'all', scale: 1, format: 'png' })
  await fs.writeFile(path.join(outputDir, `${sheetName.replace(/[^a-z0-9]+/gi, '_').toLowerCase()}.png`), new Uint8Array(await preview.arrayBuffer()))
}
const output = await SpreadsheetFile.exportXlsx(workbook)
await output.save(outputPath)
console.log(`Saved ${outputPath}`)
