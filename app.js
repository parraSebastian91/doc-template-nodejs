var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');
var express = require("express");
var app = express();
var fs = require('fs');
var path = require('path');
const Excel = require('exceljs/modern.nodejs');

app.get("/docx", async (req, res) => {
    try {
        const info = {
            //INFORME EJECUTIVO
            institucion: 'institucion',
            codigo_1: 'codigo_1',
            fecha_version: 'fecha_version',
            titulo: 'titulo',
            código_auditoria: 'código_auditoria',
            fecha_hoy: 'fecha_hoy',
            objetivo_general_auditoria: 'objetivo_general_auditoria',
            alcance_auditoria: 'alcance_auditoria',
            temas_relevantes: 'temas_relevantes',
            conclusiones_generales: {
                conclucion: 'conclucion',
                hallazgo: [
                    { detalle: 'detalle1' },
                    { detalle: 'detalle2' },
                    { detalle: 'detalle3' }
                ]
            },
            //INFORME DETALLADO
            area_proceso_auditado: 'area_proceso_auditado',
            objetivo_general_detalle: 'objetivo_general_detalle',
            objetivos_especificos: {
                objetivos: [
                    { detalle: 'objetivo1' },
                    { detalle: 'objetivo2' },
                    { detalle: 'objetivo3' }
                ]
            },
            alcance_auditoria_detalle: 'alcance_auditoria_detalle',
            oportunidad_detalle: 'oportunidad_detalle',
            equipo_trabajo: [
                {
                    nombre: 'nombre',
                    profesion: 'profesion',
                    cargo: 'cargo',
                    areas_tratadas: 'areas_tratadas'
                },
                {
                    nombre: 'nombre',
                    profesion: 'profesion',
                    cargo: 'cargo',
                    areas_tratadas: 'areas_tratadas'
                }
            ],
            metodologia_aplicada: 'metodologia_aplicada',
            limitaciones_desarrollo: 'limitaciones_desarrollo',
            resultado_detallado: [
                {
                    objetivo_especifico: 'objetivo_especifico',
                    hallazgo: 'hallazgo',
                    analisis_causa: 'analisis_causa',
                    impacto_efecto: 'impacto_efecto',
                    sugerencia: [
                        {
                            titulo: 'titulo0',
                            descripcion: 'descripcion0'
                        },
                        {
                            titulo: 'titulo1',
                            descripcion: 'descripcion0'
                        }
                    ]
                }
            ],
            acciones_revia_informe: 'acciones_revia_informe',
            compromisos: 'compromisos',
            retroalimentacion: 'retroalimentacion'
        }
        let now = new Date();
        var content = fs.readFileSync(path.resolve(__dirname, 'doc-template/Informe.docx'), 'binary');
        var zip = new PizZip(content);
        var doc = new Docxtemplater();
        doc.loadZip(zip);
        doc.setData(info);
        doc.render()
        var fileContents = doc.getZip().generate({ type: 'nodebuffer' });
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

app.get('/xlsx', async (req, res) => {
    try {
        var workbook = new Excel.Workbook();
        workbook.addWorksheet('Hoja1');
        workbook.addWorksheet('Resumen seguimiento');
        var ws = workbook.getWorksheet('Hoja1');
        ws.columns = [
            { header: 'N° Auditoría', key: 'id', width: 16 },
            { header: 'Fecha de Informe de Auditoría', key: 'fechaAuditoria', width: 16 },
            { header: 'Actividad de Auditoría', key: 'actAuditoria', width: 16 },
            { header: 'Hallazgo', key: 'DhallazgoB', width: 76 },
            { header: 'Recomendación de Auditoría', key: 'recomendacion', width: 78 },
            { header: 'Compromisos de los responsables del proceso, subproceso, etapa o materia auditada', key: 'compromiso', width: 52 },
            { header: 'Indicador de Logro del Compromiso', key: 'logro', width: 43 },
            { header: 'Plazo/Fecha Propuesta para Implementación de las medidas', key: 'plazo_propuesto', width: 16 },
            { header: 'Responsable de la implementación', key: 'estado', width: 16 },
            { header: 'Estado', key: 'estado', width: 15 },
            { header: 'Medio de verificación ', key: 'medio_verificacion', width: 37 },
            { header: 'Area', key: 'area', width: 15 },
            { header: 'OBS Auditor a Cargo', key: 'auditor', width: 45 },
        ];
        ws.getRow(1).height = 50

        // ws.getCell('B4').value = 3.14159;

        var rowValues = [];
        rowValues[1] = 4;
        rowValues[5] = 'Kyle';
        rowValues[9] = new Date();
        ws.addRow(rowValues);

        let fileName = 'testExel.xlsx';
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=' + fileName);
        console.log(fileName)
        workbook.xlsx.write(res)
            .then(function () {
                res.end();
            });
    } catch (error) {
        console.log(error)
    }
})


app.listen(3000)