const express = require('express');
const app = express();
const bodyParser = require('body-parser');
const readXlsxFile = require('read-excel-file/node');
const mysql = require('mysql');
const multer = require('multer');
const path = require('path');
const sanitizeFilename = require('sanitize-filename');

// use express static folder
app.use(express.static("./public"));

// body-parser middleware use
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Database connection
const db = mysql.createConnection({
    host: "localhost",
    user: "root",
    port: 3306,
    password: "",
    database: "exel"
});

db.connect(function (err) {
    if (err) {
        return console.error('error: ' + err.message);
    }
    console.log('Connected to the MySQL server.');
});

// Multer Upload Storage
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, path.join(__dirname, '/uploads/'));
    },
    filename: (req, file, cb) => {
        // Sanitize the filename to remove any special characters or directory traversal attempts
        const sanitizedFilename = sanitizeFilename(file.originalname);
        console.log('Sanitized Filename:', sanitizedFilename);
        cb(null, file.fieldname + "-" + Date.now() + "-" + sanitizedFilename);
    }
});

// Update multer configuration to allow any type of file
const upload = multer({ storage: storage });

// Routes
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '/index.html'));
});

app.post('/uploadfile', upload.single("uploadfile"), (req, res) => {
    console.log('Uploaded File:', req.file);
    importExcelDataToMySQL(path.join(__dirname, '/uploads/', req.file.filename));
    res.send('File uploaded successfully!');
});

function importExcelDataToMySQL(filePath) {
    console.log('File Path:', filePath);
    readXlsxFile(filePath).then((rows) => {
        rows.shift(); // Remove Header ROW

        let query = 'INSERT INTO customer (id, address, name, age) VALUES ?';
        db.query(query, [rows], (error, response) => {
            console.log(error || response);
        });
    });
}

// Create a Server
const PORT = 3000;
app.listen(PORT, () => {
    console.log(`App listening at http://localhost:${PORT}`);
});
