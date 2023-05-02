const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const signature = fs.readFileSync('path/to/image.jpg', 'base64');

const content = fs.readFileSync(
    path.resolve(__dirname, "ITD-VI-PO-01-02_Lista_de_Estudiantes_Autorizados_para_Visita.docx"),
    "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

const table = {
    rows: [
      [ '1', 'Juan Pérez', '12313442', 'Ing. En Sistemas', '8vo' ],
      [ '2', 'Jorge García', '132132312', 'Ing. Eléctrica', '2do' ],
      [ '3', 'Daniela López', '312312313', 'Ing. Industrial', '5to' ]
    ]
};

const tableIns = {
    name: 'Tabla1',
    rows: table.rows,
    options: {
      bold: true
    }
};

doc.setData({
    Tabla1: tableIns,
    signature: signature
});

doc.render({
    department: "",
    visitdate: "",
    enterpriselocation: "",
    teacher: "",
    schedule: ""
});

const data = {
    signature: signature
};

doc.render();

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "New_ITD-VI-PO-01-02_Lista_de_Estudiantes_Autorizados_para_Visita.docx"), buf);