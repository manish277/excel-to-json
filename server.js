const express = require('express')
const multer = require('multer')
const xlsx = require('xlsx')
const MongoClient = require('mongodb').MongoClient
require('dotenv').config()
const port = process.env.PORT || 3000

// Set up the express app
const app = express()
// Set up Multer for file uploads
const upload = multer({ dest: 'uploads/' })

// Set up the MongoDB connection
const url = process.env.DB_MONGOURL
// Connect to the MongoDB database
const client = new MongoClient(url, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
})
client.connect().then((con) => {
  console.log('connected to DB')
})
const coll = client
  .db(process.env.DB_NAME)
  .collection(process.env.DB_COLLECTION)

app.post('/excl-to-json', upload.single('file'), async (req, res) => {
  //read the excel file
  const workbook = xlsx.readFile(req.file.path)
  //get the first sheet
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  //get the range of data
  const range = xlsx.utils.decode_range(sheet['!ref'])
  //get the headers
  const headers = await xlsx.utils.sheet_to_json(sheet, { header: 1 })[0]
  let jsonData = []
  let totalCount
  //iterate over the rows and map the data with headers
  for (let R = range.s.r + 1; R <= range.e.r; ++R) {
    let obj = {}
    for (let C = range.s.c; C <= range.e.c; ++C) {
      let cell_address = { c: C, r: R }
      let cell_ref = xlsx.utils.encode_cell(cell_address)
      let cell_value = sheet[cell_ref]
      if (cell_value) obj[headers[C]] = cell_value.v
    }
    jsonData.push(obj)
  }
  coll.insertMany(jsonData, (error, result) => {
    if (error) throw error
    totalCount = result.insertedCount
    console.log(`Inserted ${result.insertedCount} documents into the collection`)
  })
  //send the json data as response
  res.json({
    status: 200,
    message: 'EXCEL HAS BEEN CONVERTED',
    total: 10,
    data: jsonData,
  })
})
app.listen(port, () => {
  console.log(`Listening on port ${port}`)
})
