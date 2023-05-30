const fs = require("fs");
const path = require("path");
const xlsx = require("node-xlsx");

const { INPUT_EXCEL_LOCATION, INPUT_EXCEL_NAME } = process.env;

const getExcelData = ({ sheetName }) => {
  // Read csv
  const file = xlsx.parse(
    fs.readFileSync(
      path.resolve(path.join(INPUT_EXCEL_LOCATION + "/" + INPUT_EXCEL_NAME))
    ),
    {
      defval: "",
    }
  );
  const sheetData = file.find((sheet) => sheet.name === sheetName)?.data;

  return sheetData;
};

module.exports = getExcelData;
