const ExcelJS = require("exceljs");
const XLSX = require("xlsx");

// Создаём книгу Excel в которую будем вписывать результат
const workbookWrite = new ExcelJS.Workbook();
const worksheet = workbookWrite.addWorksheet("Sheet 1");
// Создаём на первый ряд с заголовком
worksheet.addRow(["Type", "SKU", "Name", "Regular price"]);

// Загружаем файл
const base = XLSX.readFile("input/base.xlsx");
const samet = XLSX.readFile("input/samet.xlsx");
// Получаем название первого листа
const sheetNameInBase = base.SheetNames[0];
const sheetNameInSamet = samet.SheetNames[0];
// Получаем данные из первого листа
const sheetInBase = base.Sheets[sheetNameInBase];
const sheetInSamet = samet.Sheets[sheetNameInSamet];
// Преобразуем данные в формат JSON
const dataInBase = XLSX.utils.sheet_to_json(sheetInBase);
const dataInSamet = XLSX.utils.sheet_to_json(sheetInSamet);

// Задаём переменную с номером ряда
let rowNumber = 1;
// Задаём флаг для проверки наличия позиции в базе
let isAvailable = false;

// Расчёт цены
const percent = 42.86;
function countPrice(price) {
  let resultPrice = Math.ceil(price + (price * percent) / 100);
  return resultPrice;
}

// ПРОГРАММА
for (const objInSamet of dataInSamet) {
  rowNumber += 1;
  for (const objInBase of dataInBase) {
    if (objInSamet["SKU"] === objInBase["Код синхронизации"]) {
      isAvailable = true;
      worksheet.addRow([
        objInSamet["Type"],
        objInSamet["SKU"],
        objInSamet["Name"],
        countPrice(objInBase["Цена в рублях"]),
      ]);
      break;
    }
  }
  if (!isAvailable) {
    worksheet.addRow([
      objInSamet["Type"],
      objInSamet["SKU"],
      objInSamet["Name"],
      objInSamet["Regular price"],
      "Нет в базе :(",
    ]);
    worksheet.getCell(`D${rowNumber}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "F08080" },
    };
    worksheet.getCell(`E${rowNumber}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "F08080" },
    };
  }
  isAvailable = false;
}

// Записываем результат в файл
workbookWrite.xlsx.writeFile("result/result.xlsx");
