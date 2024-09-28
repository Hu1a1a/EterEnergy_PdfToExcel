const fs = require("fs");
const pdfParse = require("pdf-parse");
const ExcelJS = require("exceljs");
const path = require("path");
const exec = require("child_process").exec;
const workbookResumen = new ExcelJS.Workbook();
const worksheetResumen = workbookResumen.addWorksheet("Resumen");
const AbsPath = path.resolve();
const FolderPath = AbsPath + "\\resources\\app" + "\\archivos2\\";
const ExcelOutputPath = AbsPath + "\\resources\\app" + "\\data\\ExcelResumenFactura2.xlsx";
//const FolderPath = AbsPath + "\\archivos2\\";
//const ExcelOutputPath = AbsPath + "\\data\\ExcelResumenFactura2.xlsx";
const columna = ["Oferta", "CUPS", "Precio", "Nombre", "Compañia", "Oferta", "Repetido"]

module.exports.main2 = async () => {
  worksheetResumen.addRow(columna)
  const file = fs.readdirSync(FolderPath);
  const AllDatas = []
  for (const f of file) {
    const pdfData = await extractTextFromPDF(FolderPath + "/" + f);
    const lines = pdfData.text.split("\n");
    const extractedDatas = extractData(lines, f);
    for (const extractedData of extractedDatas) await writeDataToExcel(extractedData)
    AllDatas.push(...extractedDatas)
  }
  await paintColor(AllDatas)
  await workbookResumen.xlsx.writeFile(ExcelOutputPath);
  exec(`start "" "${ExcelOutputPath}"`);
}

const extractData = (lines, filename) => {

  const data = [];
  let i = -1
  let check = false
  lines.forEach((line, index) => {
    if (line === "xYYxBlueEnergy") check = true
    if (check) {
      const CUP_Reg = new RegExp(/^ES\d{2}.*$/);
      if (CUP_Reg.test(line)) {
        i++
        data[i] = {}
        data[i][1] = filename
        data[i][2] = line.split("xYYx")[0];
        if (line.split("xYYx").length > 1) {
          data[i][4] = line.split("xYYx")[1]
          if (line.split("xYYx")[2]) data[i][5] = line.split("xYYx")[2]
          if (line.split("xYYx")[3]) data[i][6] = line.split("xYYx")[3]
        } else {
          if (!lines[index + 1].includes("€")) {
            data[i][4] = lines[index + 1].replaceAll("xYYx", " ")
            console.log(lines[index + 2].replaceAll("xYYx", " "))
            if (!lines[index + 2].includes("€")) {
              data[i][5] = lines[index + 2].replaceAll("xYYx", " ")
              if (!lines[index + 3].includes("€")) data[i][6] = lines[index + 3].replaceAll("xYYx", " ")
              else data[i][6] = lines[index + 3].split("xYYx")[0].replaceAll("xYYx", " ")
            } else if (lines[index + 2].split("xYYx").length > 2) {
              data[i][5] = lines[index + 2].split("xYYx")[0].replaceAll("xYYx", " ")
              data[i][6] = lines[index + 2].split("xYYx")[1].replaceAll("xYYx", " ")
            }
          }
        }
      }
      if (data[i]) {
        if ((line.startsWith("2.0TD") || line.startsWith("3.0TD")) && !data[i][3]) data[i][3] = line.split("xYYx")[line.split("xYYx").length - 1]
        if ((line.includes("€") && line.includes("xYYx")) && !data[i][3]) data[i][3] = line.split("xYYx")[line.split("xYYx").length - 1]
        if (line === "€" && !data[i][3]) data[i][3] = lines[index - 1]
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
  const render_options = { normalizeWhitespace: true, disableCombineTextItems: true, };
  return pageData.getTextContent(render_options).then(function (textContent) {
    let lastY, lastX, text = "";
    for (let item of textContent.items) {
      if (lastY == item.transform[5] || !lastY) text += "xYYx" + item.str;
      else if (lastX == item.transform[4] || !lastY) text += "xYYx" + item.str;
      else text += "\n" + item.str;
      lastY = item.transform[5];
      lastX = item.transform[4];
    }
    return text;
  });
}

async function writeDataToExcel(data) {
  row = [];
  Object.keys(data).forEach((key) => row.push(...[data[key]]));
  worksheetResumen.addRow(row);
}

async function paintColor(data) {
  const dataR = encontrarRepetidos(data.map((a) => a['2'] + parseFloat(a['3'])))
  for (const d in dataR) {
    if (dataR[d]) {
      worksheetResumen.getCell(+d + 2, +Object.keys(columna)[Object.keys(columna).length - 1] + 1).value = "Repetido"
      worksheetResumen.getRow(+d + 2).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF7D7D" },
        bgColor: { argb: "FF000000" }
      }
    }
  }
}

function encontrarRepetidos(array) {
  const contador = {};
  array.forEach(item => contador[item] = (contador[item] || 0) + 1)
  return array.map(item => contador[item] > 1 ? item : null
  );
}