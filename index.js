import xlsx from 'xlsx'
import fs from 'fs'
import path from 'path'

function readFileToJson(fileName) {
  let wb = xlsx.readFile(fileName, {cellDates: null})
  let sheetName = wb.SheetNames[0]
  let ws = wb.Sheets[sheetName]
  let data = xlsx.utils.sheet_to_json(ws)
  return data
}

const dirname = path.resolve()
let dir = path.join(dirname, "Files")
let files = fs.readdirSync(dir)


let combineData =  []

files.forEach((file) => {
  let fileExtenstion = path.parse(file).ext
  if(fileExtenstion === '.xlsx' && file[0] !== '~'){
    let fullFilePath  = path.join(dirname, "Files", file)
    let data = readFileToJson(fullFilePath)
    combineData = combineData.concat(data)
  }
})


let newB = xlsx.utils.book_new()
let newWS = xlsx.utils.json_to_sheet(combineData)
xlsx.utils.book_append_sheet(newB, newWS, "sheet1")
xlsx.writeFile(newB, "combine.xlsx")
console.log("DONE")