require("dotenv").config();
const robot = require("robotjs");
const { getExcelData, sleep } = require("../utils");
const { START_ROW, INITIAL_COUNTDOWN } = process.env;

const ACTION = process.argv[2];

const isActivate = ACTION === "activate" ? true : false;

console.log("ACTION", ACTION);

// Return max lenght for each input index
const getMaxInputLenght = (index) => {
  return {
    0: 12,
    1: 4,
    2: 4,
  }[index];
};

const main = async () => {
  // TODO: wait or confirm before start;
  console.log("Sleeping, please click on first input");
  await sleep(Number(INITIAL_COUNTDOWN));
  console.log("starting...");
  const data = getExcelData({ sheetName: "status" });
  for (const row of data.slice(START_ROW, data.length)) {
    for (const columnIdx in row) {
      const value = row[columnIdx];
      if (!!value) {
        robot.typeString(value);
      }
      // Get end of row
      if (Number(columnIdx) === row.length - 1) {
        robot.keyTap("enter"); // submit and enter to
        // enter to product page
        await sleep(100);
        Array.from(Array(21)).map(async () => {
          robot.keyTap("tab");
          await sleep(10);
        });
        if (isActivate) {
          // activate input
        } else {
          // deactivating
          // solo la D y ya.
          robot.typeString(value);
        }

        // F10 para guardar.
      } else {
        if (
          String(value).trim().length < getMaxInputLenght(Number(columnIdx))
        ) {
          robot.keyTap("tab");
        }
      }
    }
    // cambio de fila
  }
};

main().catch((e) => {
  console.log(e);
  process.exit(1);
});
