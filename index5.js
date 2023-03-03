const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const content = fs.readFileSync(
    path.resolve(__dirname, "ITD-VI-PO-01-03_Oficio_de_Solicitud_de_Visita_a_Empresa.docx"),
    "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

doc.render({
    department: "",
    visitdate: "",
    enterpriselocation: "",
    teacher: "",
    schedule: "",
    number: "",
    name: "",
    id: "",
    career: "",
    semester: ""
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "New_ITD-VI-PO-01-03_Oficio_de_Solicitud_de_Visita_a_Empresa.docx"), buf);