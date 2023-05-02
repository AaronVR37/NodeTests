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
    name: "",
    activity: "",
    period: "",
    A0: "",
    A1: "",
    A2: "",
    A3: "",
    A4: "",
    B0: "",
    B1: "",
    B2: "",
    B3: "",
    B4: "",
    C0: "",
    C1: "",
    C2: "",
    C3: "",
    C4: "",
    D0: "",
    D1: "",
    D2: "",
    D3: "",
    D4: "",
    E0: "",
    E1: "",
    E2: "",
    E3: "",
    E4: "",
    F0: "",
    F1: "",
    F2: "",
    F3: "",
    F4: "",
    G0: "",
    G1: "",
    G2: "",
    G3: "",
    G4: "",
    observations: "",
    value: "",
    level: ""
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "New_ITD-AC-PO-05-02_Evaluacion_al_Desempeno_de_la_Actividad_Complementaria.docx"), buf);