const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const signature = fs.readFileSync('path/to/image.jpg', 'base64');

const content = fs.readFileSync(
    path.resolve(__dirname, "ITD-VI-PO-01-05_Carta_de_Presentacion_y_Agradecimiento.docx"),
    "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

doc.render({
    number_oficio: "",
    day: "",
    month: "",
    year: "",
    teacher: "",
    number: "",
    career: "",
    date: "",
    hours: ""
});

const data = {
    signature: signature,
};

doc.setData({
    signature: signature
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "New_ITD-VI-PO-01-05_Carta_de_Presentacion_y_Agradecimiento.docx"), buf);