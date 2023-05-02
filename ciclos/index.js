const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const signature1 = fs.readFileSync('path/to/image.jpg', 'base64');
const signature2 = fs.readFileSync('path/to/image.jpg', 'base64');
const signature3 = fs.readFileSync('path/to/image.jpg', 'base64');

const content = fs.readFileSync(
    path.resolve(__dirname, "ITD-VI-PO-01-01_Solicitud_Departamental_de_Visitas_a_Empresas.docx"),
    "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

const table = {
    rows: [
      [ '20/06/23', 'Guadalajara', 'Bimbo', '34', 'Paquito', '12:00 - 17:00', 'Ing. Informática - 4to' ],
    ]
};

const tableIns = {
    name: 'Tabla1',
    rows: table.rows,
    options: {
      bold: true
    }
};

doc.render({
    period: "",
    docdate: ""
});

const data = {
    signature1: signature1,
    signature2: signature2,
    signature3: signature3
};

doc.setData({
    Tabla1: tableIns,
    signature1: signature1,
    signature2: signature2,
    signature3: signature3
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "New_ITD-VI-PO-01-01_Solicitud_Departamental_de_Visitas_a_Empresas.docx"), buf);