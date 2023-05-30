require("dotenv").config();
const robot = require("robotjs");
const { getExcelData, sleep } = require("../utils");
const { START_ROW, INITIAL_COUNTDOWN } = process.env;

// Return max lenght for each input index
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
  const data = getExcelData({ sheetName: "load" });
  for (const row of data.slice(START_ROW, data.length)) {
    for (const columnIdx in row) {
      const value = row[columnIdx];
      // We assume that columns are:  FABRICA CONCAT | 	WAREHOUSE SKU | 	FT STYLE | 	FT COLOR	| FT SIZE	| UNITS	| Cost(WAC)	| PO
      if (!!value) {
        robot.typeString(value);
      }
      // Get end of row
      if (Number(columnIdx) === row.length - 1) {
        robot.keyTap("f9"); // submit
        robot.keyTap("f12"); // refresh
        robot.keyTap("f9"); // back to form
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
