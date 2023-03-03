const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const content = fs.readFileSync(
    path.resolve(__dirname, "ITD-VI-PO-01-02_Lista_de_Estudiantes_Autorizados_para_Visita.docx"),
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

fs.writeFileSync(path.resolve(__dirname, "New_ITD-VI-PO-01-02_Lista_de_Estudiantes_Autorizados_para_Visita.docx"), buf);