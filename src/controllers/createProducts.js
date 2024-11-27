// enter cuando estan los 3 primeros
/**
 * Aparece autotab en description.
 * Description max count 30. a partir de 30 altabea solo
 * Desde description a vendor son  8 tabs para quedar parado en vendor.
 * Vendor max count 20.
 * F9 para guardar y manda solo a la vista anterior.
 */

require("dotenv").config();
const r = require("robotjs");
const tr = require("../utils/testRobot");
const { getExcelData, sleep, range } = require("../utils");
const { START_ROW, INITIAL_COUNTDOWN } = process.env;

const robot = process.env.NODE_ENV === "development" ? tr : r;
// Return max lenght for each input index
const getMaxInputLenght = (index) => {
  return {
    0: 12,
    1: 4,
    2: 4,
    3: 30, // description
    4: 20, // vendor
    5: 10, // CUST PRICE
    6: 12, // LANDED COST
  }[index];
};

const tabsBeforeInput = (index) => {
  return {
    0: 0,
    1: 1,
    2: 2,
    3: 0, // description
    4: 8, // vendor
    5: 26, // CUST PRICE
    6: 30, // LANDED COST
  }[index];
};

const main = async () => {
  console.log("Sleeping, please click on first input");
  await sleep(Number(INITIAL_COUNTDOWN));
  console.log("starting...");
  const data = getExcelData({ sheetName: "create" });
  // A partir de este indice debería cambiar de "vista".
  const colIndexToChangeView = 3;
  for (const row of data.slice(START_ROW, data.length)) {
    let tabPos = 0;
    for (const columnIdx in row) {
      const idx = Number(columnIdx);
      const value = row[idx];
      const maxInputLenght = getMaxInputLenght(idx);
      const slicedValue = String(value)?.trim().slice(0, maxInputLenght);
      if (idx >= colIndexToChangeView) {
        // product page

        // clean field in case there is some value
        robot.keyTap("end");

        // type value
        robot.typeString(slicedValue);

        if (idx === row.length - 1) {
          // last col
          robot.keyTap("f9"); // submit save product
        } else {
          // tab necesary times before next input
          let tabs = tabsBeforeInput(idx + 1) - tabPos;
          console.log({ tabs, tabPos });
          // Current value can be equal to max and by default plattform perform a autotab so we balance
          const balance = slicedValue.length < maxInputLenght ? 0 : -1;

          range(tabs + balance).forEach(() => {
            robot.keyTap("tab");
            tabPos += 1;
          });
        }
      } else {
        // search product page

        // clean field in case there is some value
        robot.keyTap("end");

        // type value
        robot.typeString(slicedValue);

        // Tab if needed
        if (slicedValue.length < maxInputLenght) {
          robot.keyTap("tab");
          tabPos += 1;
        }

        // enter to product page
        if (idx === colIndexToChangeView - 1) {
          robot.keyTap("enter");
          tabPos = 0;
        }
      }
    }
  }
};

main().catch((e) => {
  console.log(e);
  process.exit(1);
});
