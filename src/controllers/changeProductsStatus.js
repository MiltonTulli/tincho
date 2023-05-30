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
    3: 1,
  }[index];
};

const main = async () => {
  console.log("Sleeping, please click on first input");
  await sleep(Number(INITIAL_COUNTDOWN));
  console.log("starting...");
  const data = getExcelData({ sheetName: "status" });
  for (const row of data.slice(START_ROW, data.length)) {
    for (const columnIdx in row) {
      const value = row[columnIdx].trim();
      // Get end of row - LAST COLUMN = STATUS COLUMN
      if (Number(columnIdx) === row.length - 1) {
        // enter to product page
        robot.keyTap("enter");
        // tab until status field
        Array.from(Array(21)).map(() => {
          robot.keyTap("tab");
        });
        // Clean field
        robot.keyTap("end");
        // Add value (D, '')
        robot.typeString(value);

        // save
        robot.keyTap("f10");
      } else {
        // Clean field
        robot.keyTap("end");
        // type value
        robot.typeString(value);
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
