const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const content = fs.readFileSync(
    path.resolve(__dirname, "TRANSPORTE_BIMBO_A-D_2022.docx"),
    "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

doc.render({
    date: "",
    folio: "",
    students_allowed: "",
    career: "",
    teacher: "",
    number_id: "",
    date_visit: "",
    hour_visit: "",
    enterprise: "",
    location: "",
    leave_hour: "",
    return_hour: ""
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "New_TRANSPORTE_BIMBO_A-D_2022.docx"), buf);