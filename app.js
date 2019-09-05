var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');
var express = require("express");
var app = express();
var fs = require('fs');
var path = require('path');
app.get("/", async (req, res) => {
    try {
        const info = {
            institucion: 'institucion',
            codigo_1:'codigo_1',
            fecha_version: 'fecha_version',
            titulo:'titulo',
            código_auditoria:'código_auditoria',
            fecha_hoy: 'fecha_hoy',
            objetivo_general_auditoria: 'objetivo_general_auditoria',
            alcance_auditoria: 'alcance_auditoria',
            temas_relevantes: 'temas_relevantes',
            conclusiones_generales : {
                conclucion: 'conclucion',
                hallazgo:[
                    {detalle:'detalle1'},
                    {detalle:'detalle2'},
                    {detalle:'detalle3'}
                ]
            }
        }
        let now= new Date();
        var content = fs.readFileSync(path.resolve(__dirname, 'doc-template/Informe.docx'), 'binary');
        var zip = new PizZip(content);
        var doc = new Docxtemplater();
        doc.loadZip(zip);
        doc.setData(info);
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