require("dotenv").config();
const fs = require("fs");
const path = require("path");
const xlsx = require("node-xlsx");
const robot = require("robotjs");

const { INPUT_EXCEL_LOCATION, INPUT_EXCEL_NAME, SHEET_NAME, START_ROW } =
  process.env;

const getData = () => {
  // Read csv
  const file = xlsx.parse(
    fs.readFileSync(
      path.resolve(path.join(INPUT_EXCEL_LOCATION + "/" + INPUT_EXCEL_NAME))
    )
  );
  const sheetData = file.find((sheet) => sheet.name === SHEET_NAME)?.data;

  return sheetData;
};

function sleep(s) {
  return new Promise((resolve) => setTimeout(resolve, s * 1000));
}

const main = async () => {
  // TODO: wait or confirm before start;
  console.log("Sleeping, please click on first input");
  await sleep(2);
  console.log("starting...");
  const data = getData();
  const fields = [3, 4, 5, 6, 7];
  for (const row of data.slice(START_ROW, data.length)) {
    for (const columnIdx in row) {
      const value = row[columnIdx];
      console.log("value", value);
      // We assume that columns are:  FABRICA CONCAT | 	WAREHOUSE SKU | 	FT STYLE | 	FT COLOR	| FT SIZE	| UNITS	| Cost(WAC)	| PO
      if (fields.includes(Number(columnIdx))) {
        robot.typeString(value);
        robot.keyTap("tab");
      }
      // Get end of row
      if (Number(columnIdx) === row.length - 1) {
        // robot.keyTap("f9");
        robot.keyTap("enter");
      }
    }
  }
};

main().catch((e) => {
  console.log(e);
  process.exit(1);
});
