const fs = require("fs");
const pdfParse = require("pdf-parse");
const ExcelJS = require("exceljs");

const Parameter = {
  NumeroFactura: { Selector: "Nº factura", Separator: ":", Value: "" },
  NombreCliente: { Selector: "Titular del contrato", Separator: ":", Value: "" },
  NIF: { Selector: "NIF", Separator: ":", Value: "" },
  CUPS: { Selector: "CUPS", Separator: ":", Value: "" },
  TarifaAcceso: { Selector: "Peaje de transporte y distribución", Separator: ":", Value: "" },
  Fecha: { Selector: "Periodo de facturación", Separator: ":", Value: "" },
  FechaFactura: { Selector: "Fecha emisión factura", Separator: ":", Value: "" },
  ImporteTotalFactura: { Selector: "Total", Separator: "Total", Value: "", Check: "€" },
  ImportePotencia: { Selector: "Potencia", Separator: "Potencia", Value: "", Check: "€" },
  ImporteEnergia: { Selector: "Energía", Separator: "Energía", Value: "", Check: "€" },
  ImporteOtros: { Selector: "Otros", Separator: "Otros", Value: "", Check: "€" },
  ImporteImpuesto: { Selector: "Impuestos", Separator: "Impuestos", Value: "", Check: "€" },
  ConsumoPunta: { Selector: "Punta", Separator: "Punta", Value: "", Check: "." },
  ConsumoLlano: { Selector: "Llano", Separator: "Llano", Value: "", Check: "." },
  ConsumoValle: { Selector: "Valle", Separator: "Valle", Value: "", Check: "." },
};
const extractData = (lines) => {
  const data = {};
  lines.forEach((line) => {
    for (const item of Object.keys(Parameter)) {
      if (line.includes(Parameter[item].Selector)) {
        if (Parameter[item].Check) {
          if (line.split(Parameter[item].Separator)[1]?.trim().includes(Parameter[item].Check)) {
            data[item] = line.split(Parameter[item].Separator)[1]?.trim();
          }
        } else {
          data[item] = line.split(Parameter[item].Separator)[1]?.trim();
        }
        console.log(item, ":", data[item]);
      }
    }
  });
  return data;
};
async function extractTextFromPDF(pdfPath) {
  const dataBuffer = fs.readFileSync(pdfPath);
  return await pdfParse(dataBuffer);
}
async function writeDataToExcel(data, excelPath) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Factura Endesa");
  worksheet.addRow(["Campo", "Valor"]);
  Object.keys(data).forEach((key) => {
    if (typeof data[key] === "object") {
      Object.keys(data[key]).forEach((subKey) => {
        worksheet.addRow([`${key} ${subKey}`, data[key][subKey]]);
      });
    } else {
      worksheet.addRow([key, data[key]]);
    }
  });

  await workbook.xlsx.writeFile(excelPath);
}

async function pdfToExcel(pdfPath, excelPath) {
  const pdfData = await extractTextFromPDF(pdfPath);
  const lines = pdfData.text.split("\n");
  const extractedData = extractData(lines);
  await writeDataToExcel(extractedData, excelPath);
  console.log(`Archivo Excel creado: ${excelPath}`);
}

const pdfPath = "./archivos/endesa.pdf";
const excelPath = "./archivos/facturaEndesa.xlsx";

pdfToExcel(pdfPath, excelPath).catch(console.error);
