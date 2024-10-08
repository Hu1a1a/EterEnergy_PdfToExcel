const fs = require("fs");
const pdfParse = require("pdf-parse");
const ExcelJS = require("exceljs");
const path = require("path");
const exec = require("child_process").exec;
const AbsPath = path.resolve();
//const ParameterPath = AbsPath + "\\data\\PdfToExcel_Parametros.xlsx";
//const FolderPath = AbsPath + "\\archivos\\";
//const ExcelOutputPath = AbsPath + "\\data\\ExcelResumenFactura.xlsx";
const ParameterPath = AbsPath + "\\resources\\app" + "\\data\\PdfToExcel_Parametros.xlsx";
const FolderPath = AbsPath + "\\resources\\app" + "\\archivos\\";
const ExcelOutputPath = AbsPath + "\\resources\\app" + "\\data\\ExcelResumenFactura.xlsx";

module.exports.main = async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(ParameterPath);
  const parameterrow = [];
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
      const lines = pdfData.text.split("\n");
      const extractedData = extractData(lines, w.getSheetValues(), i);
      await writeDataToExcel(extractedData, w.name, f);
    }
  }
  await workbookResumen.xlsx.writeFile(ExcelOutputPath);
  exec(`start "" "${ExcelOutputPath}"`);
}

const extractData = (lines, parameter, indexx) => {
  const data = {};
  for (let i = 1; i < parameter.length - 1; i++) data[parameter[i + 1][1]] = null;
  lines.forEach((line, index) => {
    for (let i = 1; i < parameter.length - 1; i++) {
      if (parameter[i + 1][2]) {
        for (const match of parameter[i + 1][2].split("&&")) {
          if (line.includes(match)) {
            data[parameter[i + 1][1]] = line.split(match)[1]?.trim();
            if (parameter[i + 1][8]) {
              data[parameter[i + 1][1]] = lines[index + parameter[i + 1][8]]?.trim();
            }
            if (parameter[i + 1][3] && !data[parameter[i + 1][1]].includes(parameter[i + 1][3])) {
              data[parameter[i + 1][1]] = null;
            } else {
              if (parameter[i + 1][4]) {
                if (parameter[i + 1][5]) {
                  if (parameter[i + 1][5] >= 0) {
                    if (data[parameter[i + 1][1]].split(parameter[i + 1][4])[parameter[i + 1][5]]?.trim())
                      data[parameter[i + 1][1]] = data[parameter[i + 1][1]].split(parameter[i + 1][4])[parameter[i + 1][5]]?.trim();
                    else data[parameter[i + 1][1]] = data[parameter[i + 1][1]].split(parameter[i + 1][4])[parameter[i + 1][5] + 1]?.trim();
                  } else {
                    if (data[parameter[i + 1][1]].split(parameter[i + 1][4])[data[parameter[i + 1][1]].split(parameter[i + 1][4]).length + parameter[i + 1][5]]?.trim())
                      data[parameter[i + 1][1]] = data[parameter[i + 1][1]].split(parameter[i + 1][4])[data[parameter[i + 1][1]].split(parameter[i + 1][4]).length + parameter[i + 1][5]]?.trim();
                    else data[parameter[i + 1][1]] = data[parameter[i + 1][1]].split(parameter[i + 1][4])[data[parameter[i + 1][1]].split(parameter[i + 1][4]).length + parameter[i + 1][5] - 1]?.trim();
                  }
                }
                if (parameter[i + 1][6]) {
                  if (data[parameter[i + 1][1]].split(parameter[i + 1][6])[parameter[i + 1][7]]?.trim())
                    data[parameter[i + 1][1]] = data[parameter[i + 1][1]].split(parameter[i + 1][6])[parameter[i + 1][7]]?.trim();
                  else data[parameter[i + 1][1]] = data[parameter[i + 1][1]].split(parameter[i + 1][6])[parameter[i + 1][7] + 1]?.trim();
                }
              }
              if (parameter[i + 1][9]) {
                if (parameter[i + 1][9] > 0) {
                  data[parameter[i + 1][1]] = data[parameter[i + 1][1]].substring(parameter[i + 1][9])?.trim();
                } else {
                  data[parameter[i + 1][1]] = data[parameter[i + 1][1]].slice(parameter[i + 1][9])?.trim();
                }
              }
            }
          }
        }
        data[parameter[i + 1][1]] = data[parameter[i + 1][1]]?.trim().replaceAll("xYYx", "");
      } else if (parameter[i + 1][10]) {
        data[parameter[i + 1][1]] = parameter[i + 1][10]?.trim();
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

async function writeDataToExcel(data, name, f) {
  row = [name + "_" + f];
  Object.keys(data).forEach((key) => row.push(...[data[key]]));
  worksheetResumen.addRow(row);
}
