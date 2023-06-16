// enter cuando estan los 3 primeros
/**
 * Aparece autotab en description.
 * Description max count 30. a partir de 30 altabea solo
 * Desde description a vendor son  8 tabs para quedar parado en vendor.
 * Vendor max count 20.
 * F9 para guardar y manda solo a la vista anterior.
 */

require("dotenv").config();
const robot = require("robotjs");
const { getExcelData, sleep, range } = require("../utils");
const { START_ROW, INITIAL_COUNTDOWN } = process.env;

// Return max lenght for each input index
const getMaxInputLenght = (index) => {
  return {
    0: 12,
    1: 4,
    2: 4,
    3: 30, // description
    4: 20, // vendor
  }[index];
};

const tabsBeforeInput = (index) => {
  return {
    0: 0,
    1: 1,
    2: 2,
    3: 0, // description
    4: 8, // vendor
  }[index];
};

const main = async () => {
  console.log("Sleeping, please click on first input");
  await sleep(Number(INITIAL_COUNTDOWN));
  console.log("starting...");
  const data = getExcelData({ sheetName: "create" });
  const colIndexToChangeView = 3;
  for (const row of data.slice(START_ROW, data.length)) {
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
          let tabs = tabsBeforeInput(idx + 1);
          // Current value can be equal to max and by default plattform perform a autotab so we balance
          const balance = slicedValue.length < maxInputLenght ? 0 : -1;

          range(tabs + balance).forEach(() => robot.keyTap("tab"));
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
        }

        // enter to product page
        if (idx === colIndexToChangeView - 1) {
          robot.keyTap("enter");
        }
      }
    }
  }
};

main().catch((e) => {
  console.log(e);
  process.exit(1);
});
