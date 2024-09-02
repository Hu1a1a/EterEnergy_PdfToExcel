const fs = require("fs");
const pdfParse = require("pdf-parse");
const ExcelJS = require("exceljs");
const path = require("path");
const exec = require("child_process").exec;

const AbsPath = path.resolve();
const FolderPath = AbsPath + "\\archivos2\\";
const ExcelOutputPath = AbsPath + "\\ExcelResumenFactura.xlsx";
console.log(
  `
  Se esta procediendo a la transformación de PDF a Excel!
  
  Dedicado a EterEnergy S.L.
  Autor: YK Web Studio - Yang Ye
  Contacto: https://hu1a1a.github.io/YK-Web-Studio/

`
);
main();
async function main() {
  const file = fs.readdirSync(FolderPath);
  for (const f of file) {
    const pdfData = await extractTextFromPDF(FolderPath + "/" + f);
    const lines = pdfData.text.split("\n");
    if (f == "8010.pdf") { console.log(lines) }
    const extractedDatas = extractData(lines);
    for (const extractedData of extractedDatas) await writeDataToExcel(extractedData, f);
  }
  await workbookResumen.xlsx.writeFile(ExcelOutputPath);
  exec(`start "" "${ExcelOutputPath}"`);
}
const extractData = (lines) => {
  const data = [];
  let i = -1
  let check = false
  lines.forEach((line, index) => {
    if (line === "xYYxBlueEnergy") check = true
    if (check) {
      if (line.startsWith("ES")) {
        i++
        data[i] = {}
        data[i][2] = line.split("xYYx")[0];
      }
      if ((line.startsWith("2.0TD") || line.startsWith("3.0TD")) && !data[i][3]) {
        data[i][3] = line.split("xYYx")[line.split("xYYx").length - 1]
      }
      if ((line.includes("€") && line.includes("xYYx")) && !data[i][4]) {
        data[i][4] = line.split("xYYx")[line.split("xYYx").length - 1]
      }
      if (line === "€" && !data[i][5]) {
        data[i][5] = lines[index - 1]
      }
    }
  });
  return data;
};

async function extractTextFromPDF(pdfPath) {
  const dataBuffer = fs.readFileSync(pdfPath);
  return await pdfParse(dataBuffer, { pagerender: render_page, version: "v2.0.550" });
}

function render_page(pageData) {
  let render_options = {
    normalizeWhitespace: true,
    disableCombineTextItems: true,
  };
  return pageData.getTextContent(render_options).then(function (textContent) {
    let lastY,
      text = "";
    for (let item of textContent.items) {
      if (lastY == item.transform[5] || !lastY) {
        text += "xYYx" + item.str;
      } else {
        text += "\n" + item.str;
      }
      lastY = item.transform[5];
    }
    return text;
  });
}

const workbookResumen = new ExcelJS.Workbook();
const worksheetResumen = workbookResumen.addWorksheet("Resumen");

async function writeDataToExcel(data, f) {
  row = [f];
  Object.keys(data).forEach((key) => row.push(...[data[key]]));
  worksheetResumen.addRow(row);
}
