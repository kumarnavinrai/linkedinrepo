const express = require('express')
const fs = require('fs')
const path = require('path')
const bodyparser = require('body-parser')
const readXlsxFile = require('read-excel-file/node')
const mysql = require('mysql')
const multer = require('multer')
const app = express()
app.use(express.static('./public'))
app.use(bodyparser.json())
app.use(
  bodyparser.urlencoded({
    extended: true,
  }),
)
const db = mysql.createConnection({
  host: 'localhost',
  user: 'root',
  password: 'Aa@12345',
  database: 'gbotdatabase',
})
db.connect(function (err) {
  if (err) {
    return console.error('error: ' + err.message)
  }
  console.log('Database connected.')
})
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, __dirname + '/uploads/')
  },
  filename: (req, file, cb) => {
    cb(null, file.fieldname + '-' + Date.now() + '-' + file.originalname)
  },
})
const uploadFile = multer({ storage: storage })
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html')
})
app.post('/import-excel', uploadFile.single('import-excel'), (req, res) => {
  importFileToDb(__dirname + '/uploads/' + req.file.filename)
  //console.log(res)
})
function importFileToDb(exFile) {
  readXlsxFile(exFile).then((rows) => {
    rows.shift()
    
    var config = {
        user: 'root',
        password: 'Aa@12345',
        host: 'localhost',
        database: 'gbotdatabase',
        port: 3306
    };
    var con = mysql.createConnection({
		user: 'root',
		password: 'Aa@12345',
		host: 'localhost',
		database: 'gbotdatabase'
	});
    var insrows = [];
    var firstname = '';
    var lastname = '';
    var month = '';
    var day = '';
    var year = '';
    var gender = '';

    rows.map((item) => {
      firstname = item[0]?item[0]:'';
      lastname = item[1]?item[1]:'';
      month = item[4]?item[4].toString():'';
      day = item[5]?item[5].toString():'';
      year = item[6]?item[6].toString():'';
      gender = item[8]?item[8].toString():'';

      insrows.push([firstname,lastname,month,day,year,gender])
    });
    insrows = insrows.splice(4101,2000);
    console.log(insrows)
    con.connect(config, function (err){
		if(err) console.log(err);
		con.query('INSERT INTO createemails (firstname, lastname, month, day, year, gender) VALUES ?', [insrows], function (errq, recordset) {
			if(errq) console.log(errq);
		    console.log(recordset);
		});
	});
    // db.connect((error) => {
    //   if (error) {
    //     console.error(error)
    //   } else {
    //     let query = 'INSERT INTO user (id, name, email) VALUES ?'
    //     db.query(query, [rows], (error, response) => {
    //       console.log(error || response)
    //     })
    //   }
    // })
  })
}
let nodeServer = app.listen(4000, function () {
  let port = nodeServer.address().port
  let host = nodeServer.address().address
  console.log('App working on: ', host, port)
})