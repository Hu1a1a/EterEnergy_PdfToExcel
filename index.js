const fs = require("fs");
const pdfParse = require("pdf-parse");
const ExcelJS = require("exceljs");

const ParameterPath = "./archivos/PdfToExcel_Parametros.xlsx";
const ExcelOutputPath = "./archivos/ExcelResumenFactura.xlsx";
const FolderPath = "./archivos/";

main();
async function main() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(ParameterPath);
  const parameterrow = ["archivo"];
  workbook
    .getWorksheet(1)
    .getColumn(1)
    .eachCell((value) => parameterrow.push(...[value.value.toString()]));
  worksheetResumen.addRow(parameterrow);
  i = 0;
  for (const w of workbook.worksheets) {
    const file = fs.readdirSync(FolderPath + w.name);
    for (const f of file) {
      i++;
      const pdfData = await extractTextFromPDF(FolderPath + w.name + "/" + f);
      //if (w.name === "nexus") console.log(pdfData);
      const lines = pdfData.text.split("\n");
      const extractedData = extractData(lines, w.getSheetValues(), i);
      await writeDataToExcel(extractedData, w.name, f);
    }
  }
  await workbookResumen.xlsx.writeFile(ExcelOutputPath);
}

const extractData = (lines, parameter, indexx) => {
  const data = {};
  for (let i = 1; i < parameter.length - 1; i++) data[parameter[i + 1][1]] = "";
  lines.forEach((line, index) => {
    for (let i = 1; i < parameter.length - 1; i++) {
      if (parameter[i + 1][2]) {
        if (line.includes(parameter[i + 1][2])) {
          data[parameter[i + 1][1]] = line.split(parameter[i + 1][2])[1]?.trim();
          if (parameter[i + 1][8]) {
            data[parameter[i + 1][1]] = lines[index + parameter[i + 1][8]]?.trim();
          }
          if (parameter[i + 1][3] && !data[parameter[i + 1][1]].includes(parameter[i + 1][3])) {
            data[parameter[i + 1][1]] = "";
          }
          if (parameter[i + 1][4]) {
            data[parameter[i + 1][1]] = data[parameter[i + 1][1]].split(parameter[i + 1][4])[parameter[i + 1][5]]?.trim();
            if (parameter[i + 1][6]) {
              data[parameter[i + 1][1]] = data[parameter[i + 1][1]].split(parameter[i + 1][6])[parameter[i + 1][7]]?.trim();
            }
          }
          if (parameter[i + 1][9]) {
            if (parameter[i + 1][9] > 0) {
              data[parameter[i + 1][1]] = data[parameter[i + 1][1]].substring(parameter[i + 1][9]);
            } else {
              data[parameter[i + 1][1]] = data[parameter[i + 1][1]].slice(parameter[i + 1][9]);
            }
          }
          if (parameter[i + 1][3] === "€") {
            data[parameter[i + 1][1]] = parseFloat(data[parameter[i + 1][1]]);
          }
        }
      } else if (parameter[i + 1][10]) {
        data[parameter[i + 1][1]] = parameter[i + 1][10];
      } else if (parameter[i + 1][11]) {
        data[parameter[i + 1][1]] = { formula: parameter[i + 1][11].replaceAll("xRx", indexx + 1) };
      } else {
        data[parameter[i + 1][1]] = null;
      }
    }
  });
  return data;
};

async function extractTextFromPDF(pdfPath) {
  const dataBuffer = fs.readFileSync(pdfPath);
  return await pdfParse(dataBuffer);
}

const workbookResumen = new ExcelJS.Workbook();
const worksheetResumen = workbookResumen.addWorksheet("Resumen");

async function writeDataToExcel(data, name, f) {
  row = [name, f];
  Object.keys(data).forEach((key) => row.push(...[data[key]]));
  worksheetResumen.addRow(row);
}
