const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const content = fs.readFileSync( //En esta parte el código reconoce el documento que se estará empleando para ser modificado
    path.resolve(__dirname, "ITD-VI-PO-01-01_Solicitud_Departamental_de_Visitas_a_Empresas.docx"),
    "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

doc.render({ //Aquí se están definiendo los valores de los campos que serán modificados
    period: "Hello",
    docdate: "12/03/23",
    number: "1",
    enterprise_city: "QACS - Durango",
    objective: "Yes",
    date: "18/03/23",
    career: "Ing. En Sistemas",
    numberstudents: "18",
    semester: "Séptimo"
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

//Finalmente usa la infomación proporcionada para usarse en los campos y genera el nuevo documento con los cambios realizados
fs.writeFileSync(path.resolve(__dirname, "New_ITD-VI-PO-01-01_Solicitud_Departamental_de_Visitas_a_Empresas.docx"), buf);