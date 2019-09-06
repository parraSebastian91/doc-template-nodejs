var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');
var express = require("express");
var app = express();
var fs = require('fs');
var path = require('path');
app.get("/", async (req, res) => {
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
                    hallazgos: {
                        descripcion:'',
                        hallazgo:[
                            {detalle:'detalle1'},
                            {detalle:'detalle2'},
                            {detalle:'detalle3'},
                            {detalle:'detalle4'}
                        ]
                    },
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
                },
                {
                    objetivo_especifico: 'objetivo_especificoB',
                    hallazgos: {
                        descripcion:'',
                        hallazgo:[
                            {detalle:'detalle1B'},
                            {detalle:'detalle2B'},
                            {detalle:'detalle3B'},
                            {detalle:'detalle4B'}
                        ]
                    },
                    analisis_causa: 'analisis_causaB',
                    impacto_efecto: 'impacto_efectoB',
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
            acciones_revia_informe:'acciones_revia_informe',
            compromisos:'compromisos',
            retroalimentacion:'retroalimentacion'
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


app.listen(3000)