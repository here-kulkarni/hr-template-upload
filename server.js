const express = require("express");
const excelToJson = require('convert-excel-to-json');
var cors = require('cors')
const fs = require('fs');
let multer = require('multer');
const { parseSheet } = require('./parse')
const fsPromises = fs.promises
const AdmZip = require('adm-zip');
const file = new AdmZip();

const PORT = process.env.PORT || 3001;

var upload = multer()
const app = express();
app.use(cors())

app.get('/', function (req, res) {
    res.send('I am working')
})

app.get("/files", async (req, res) => {
    file.addLocalFolder('./downloads');
    fs.writeFileSync('./downloads/output.zip', file.toBuffer());
    res.download('./downloads/output.zip')
})

app.post('/import', upload.any(), async function (req, res) {
    const directory = process.cwd() + "/downloads";
    if (fs.existsSync(directory)) {
        if (fs.readdirSync(directory).length !== 0) {
            await fsPromises.rmdir(directory, {
                recursive: true
            })
            fs.mkdirSync(directory);
        }

    } else {
        fs.mkdirSync(directory);
    }
    const form = req.files
    const result = excelToJson({
        source: form[0].buffer
    });
    parseSheet(result.Sheet1, form[1]);
    res.send('Successfully file converted')

})


app.listen(PORT, () => {
    console.log(`Server listening on ${PORT}`);
});