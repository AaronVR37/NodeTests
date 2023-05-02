const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const signature1 = fs.readFileSync('path/to/image.jpg', 'base64');
const signature2 = fs.readFileSync('path/to/image.jpg', 'base64');

const content = fs.readFileSync(
    path.resolve(__dirname, "ITD-VI-PO-01-06_Reporte_de_Resultados_e_Incidencias_en_Visitas_a_Empresas.docx"),
    "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

doc.render({
    date: "",
    teacher: "",
    career: "",
    enterprise: "",
    visitdate: "",
    gender: "",
    visithour: "",
    class: "",
    objectives: "",
    goals: "",
    units: "",
    issues: ""
});

const data = {
    signature1: signature1,
    signature2: signature2
};

doc.setData({
    signature1: signature1,
    signature2: signature2
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "New_ITD-VI-PO-01-06_Reporte_de_Resultados_e_Incidencias_en_Visitas_a_Empresas.docx"), buf);