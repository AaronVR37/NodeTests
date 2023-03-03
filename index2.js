const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const content = fs.readFileSync(
    path.resolve(__dirname, "ITD-AC-PO-05-02_Evaluacion_al_Desempeno_de_la_Actividad_Complementaria.docx"),
    "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

doc.render({
    name: "José Luis Rodríguez Chávez",
    activity: "Investigación",
    period: "30 de Marzo - 5 de Abril"
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "New_ITD-AC-PO-05-02_Evaluacion_al_Desempeno_de_la_Actividad_Complementaria.docx"), buf);