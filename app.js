var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');
var express = require("express");
var app = express();
var fs = require('fs');
var path = require('path');
app.get("/", async (req, res) => {
    try {
        let now= new Date();
        var content = fs.readFileSync(path.resolve(__dirname, 'doc-template/Informe.docx'), 'binary');
        var zip = new PizZip(content);
        var doc = new Docxtemplater();
        doc.loadZip(zip);
        doc.setData({
            institucion: 'Defensa Nacional',
            fecha: now,
            objetivo_general: '0652455478',
            description: 'New Website'
        });
        doc.render()
        var fileContents = doc.getZip().generate({type: 'nodebuffer'});
        res.writeHead(200, {
            'Content-Type': "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            'Content-disposition': 'attachment;filename=test.docx',
            'Content-Length': fileContents.length
        });
        res.end(fileContents);

    } catch (error) {
        console.log(error)
    }
})


app.listen(3000)