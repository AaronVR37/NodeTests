const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const signature = fs.readFileSync('path/to/image.jpg', 'base64');
const signature2 = fs.readFileSync('path/to/image.jpg', 'base64');
const seal = fs.readFileSync('path/to/image.jpg', 'base64');

const content = fs.readFileSync(
    path.resolve(__dirname, "ITD-AC-PO-05-03_Constancia_de_cumplimiento_de_actividad_complementaria.docx"),
    "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

doc.render({
    nameboss: "",
    namefirst: "",
    name: "",
    id: "",
    namecareer: "",
    level: "",
    number: "",
    period: "",
    credits: "",
    place: "",
    dateday: "",
    month: "",
    year: ""
});

const data = {
    signature: signature,
    signature2: signature2,
    seal: seal
};

doc.setData({
    signature: signature,
    signature2: signature2,
    seal: seal
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "New_ITD-AC-PO-05-03_Constancia_de_cumplimiento_de_actividad_complementaria.docx"), buf);