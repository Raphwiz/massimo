import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool'

const here = path.dirname(fileURLToPath(import.meta.url))
const backendRoot = path.dirname(here)
const args = Object.fromEntries(process.argv.slice(2).reduce((pairs, item, index, values) => item.startsWith('--') ? [...pairs, [item.slice(2), values[index + 1]]] : pairs, []))
const className = String(args.className || '').trim()
const date = String(args.date || '').trim()
const outputPath = String(args.output || '').trim()
if (!className || !/^\d{4}-\d{2}-\d{2}$/.test(date) || !outputPath) throw new Error('A class name, YYYY-MM-DD date, and output path are required.')

const sourcePath = args.source ? path.resolve(String(args.source)) : path.join(backendRoot, 'data', 'school-data.json')
const data = JSON.parse(await fs.readFile(sourcePath, 'utf8'))
const admissionById = new Map((data.admissions || []).map(item => [item.id, item]))
const records = (data.attendanceRecords || []).filter(item => item.className === className && item.date === date).sort((a, b) => a.learnerName.localeCompare(b.learnerName))
if (!records.length) throw new Error('No saved attendance records are available for this class and date.')

const workbook = Workbook.create()
const sheet = workbook.worksheets.add('Attendance Register')
sheet.showGridLines = false
const navy = '#173B6B'
const blue = '#2B72C8'
const gold = '#D69B16'
const paleBlue = '#EDF5FF'
const paleGold = '#FFF8E8'
const border = '#DCE5EF'

sheet.mergeCells('A1:F1')
sheet.getRange('A1').values = [["St. Dorothy's School Massimo — Attendance Register"]]
sheet.getRange('A1:F1').format = { fill: navy, font: { bold: true, color: '#FFFFFF', size: 16 }, horizontalAlignment: 'left', verticalAlignment: 'center' }
sheet.getRange('A1:F1').format.rowHeight = 30
sheet.mergeCells('A2:F2')
sheet.getRange('A2').values = [[`Class: ${className}  |  Date: ${date}  |  Prepared from the saved teacher register`]]
sheet.getRange('A2:F2').format = { fill: paleBlue, font: { color: navy, italic: true }, verticalAlignment: 'center' }
sheet.getRange('A2:F2').format.rowHeight = 22
sheet.getRange('A4:F4').values = [['Registered learners', 'Present / Late', 'Absent', 'Excused', 'Attendance rate', 'Exported at']]
sheet.getRange('A4:F4').format = { fill: paleGold, font: { bold: true, color: navy }, horizontalAlignment: 'center', verticalAlignment: 'center', wrapText: true }
sheet.getRange('A5').formulas = [[`=COUNTA(C8:C${records.length + 7})`]]
sheet.getRange('B5').formulas = [[`=COUNTIF(D8:D${records.length + 7},"Present")+COUNTIF(D8:D${records.length + 7},"Late")`]]
sheet.getRange('C5').formulas = [[`=COUNTIF(D8:D${records.length + 7},"Absent")`]]
sheet.getRange('D5').formulas = [[`=COUNTIF(D8:D${records.length + 7},"Excused")`]]
sheet.getRange('E5').formulas = [[`=IF(A5=0,0,B5/A5)`]]
sheet.getRange('F5').values = [[new Date()]]
sheet.getRange('A5:F5').format = { fill: '#FFFFFF', font: { bold: true, color: navy }, horizontalAlignment: 'center', verticalAlignment: 'center', borders: { preset: 'outside', style: 'thin', color: border } }
sheet.getRange('E5').format.numberFormat = '0%'
sheet.getRange('F5').format.numberFormat = 'yyyy-mm-dd hh:mm'
sheet.getRange('A7:F7').values = [['No.', 'Admission number', 'Learner name', 'Attendance status', 'Reason / note', 'Marked at']]
sheet.getRange('A7:F7').format = { fill: blue, font: { bold: true, color: '#FFFFFF' }, horizontalAlignment: 'center', verticalAlignment: 'center', wrapText: true }
sheet.getRange('A7:F7').format.rowHeight = 24
sheet.getRange(`A8:F${records.length + 7}`).values = records.map((record, index) => [
  index + 1,
  admissionById.get(record.admissionId)?.admissionNumber || '',
  record.learnerName,
  record.status,
  record.reason || '',
  new Date(record.markedAt)
])
sheet.getRange(`A8:F${records.length + 7}`).format = { verticalAlignment: 'center', wrapText: true, borders: { preset: 'inside', style: 'thin', color: border } }
sheet.getRange(`A8:B${records.length + 7}`).format.horizontalAlignment = 'center'
sheet.getRange(`D8:D${records.length + 7}`).format.horizontalAlignment = 'center'
sheet.getRange(`F8:F${records.length + 7}`).format.numberFormat = 'yyyy-mm-dd hh:mm'
sheet.getRange(`D8:D${records.length + 7}`).conditionalFormats.add('containsText', { text: 'Present', format: { fill: '#EAF7EE', font: { color: '#216E3A', bold: true } } })
sheet.getRange(`D8:D${records.length + 7}`).conditionalFormats.add('containsText', { text: 'Late', format: { fill: '#FFF4D8', font: { color: '#9A6500', bold: true } } })
sheet.getRange(`D8:D${records.length + 7}`).conditionalFormats.add('containsText', { text: 'Absent', format: { fill: '#FDECEC', font: { color: '#B42318', bold: true } } })
sheet.getRange(`D8:D${records.length + 7}`).conditionalFormats.add('containsText', { text: 'Excused', format: { fill: '#F0EDFF', font: { color: '#5940A6', bold: true } } })
sheet.tables.add(`A7:F${records.length + 7}`, true, 'AttendanceRegister')
sheet.freezePanes.freezeRows(7)
for (const [column, width] of Object.entries({ A: 8, B: 22, C: 30, D: 20, E: 38, F: 23 })) sheet.getRange(`${column}:${column}`).format.columnWidth = width
sheet.getRange(`A8:F${records.length + 7}`).format.autofitRows()
sheet.getRange(`A1:F${records.length + 7}`).format.font = { name: 'Aptos', size: 10 }
sheet.getRange('A1:F1').format.font = { name: 'Aptos Display', size: 16, bold: true, color: '#FFFFFF' }

await fs.mkdir(path.dirname(outputPath), { recursive: true })
const output = await SpreadsheetFile.exportXlsx(workbook)
await output.save(outputPath)
