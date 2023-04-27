require("dotenv").config();
const fs = require("fs");
const path = require("path");
const xlsx = require("node-xlsx");
const robot = require("robotjs");

const {
  INPUT_EXCEL_LOCATION,
  INPUT_EXCEL_NAME,
  SHEET_NAME,
  START_ROW,
  INITIAL_COUNTDOWN,
} = process.env;

const getData = () => {
  // Read csv
  const file = xlsx.parse(
    fs.readFileSync(
      path.resolve(path.join(INPUT_EXCEL_LOCATION + "/" + INPUT_EXCEL_NAME))
    ),
    {
      defval: "",
    }
  );
  const sheetData = file.find((sheet) => sheet.name === SHEET_NAME)?.data;

  return sheetData;
};

function sleep(s) {
  return new Promise((resolve) => setTimeout(resolve, s * 1000));
}

const getMaxInputLenght = (index) => {
  return {
    0: 12,
    1: 4,
    2: 4,
    3: 4,
    4: 20,
    5: 7,
    6: 11,
  }[index];
};

const main = async () => {
  // TODO: wait or confirm before start;
  console.log("Sleeping, please click on first input");
  await sleep(Number(INITIAL_COUNTDOWN));
  console.log("starting...");
  const data = getData();
  for (const row of data.slice(START_ROW, data.length)) {
    for (const columnIdx in row) {
      const value = row[columnIdx];
      // We assume that columns are:  FABRICA CONCAT | 	WAREHOUSE SKU | 	FT STYLE | 	FT COLOR	| FT SIZE	| UNITS	| Cost(WAC)	| PO
      if (!!value) {
        2;
        robot.typeString(value);
      }
      // Get end of row
      if (Number(columnIdx) === row.length - 1) {
        robot.keyTap("f9"); // submit
        robot.keyTap("f12"); // refresh
        robot.keyTap("f9"); // back to form
        // robot.keyTap("enter");
      } else {
        if (
          String(value).trim().length < getMaxInputLenght(Number(columnIdx))
        ) {
          robot.keyTap("tab");
        }
      }
    }
  }
};

main().catch((e) => {
  console.log(e);
  process.exit(1);
});
