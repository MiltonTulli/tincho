require("dotenv").config();
const robot = require("robotjs");
const { getExcelData, sleep } = require("../utils");
const { START_ROW, INITIAL_COUNTDOWN } = process.env;

const main = async () => {
  // TODO: wait or confirm before start;
  console.log("Sleeping, please click on first input");
  await sleep(Number(INITIAL_COUNTDOWN));
  console.log("starting...");
  Array.from(Array(21)).map(async () => {
    robot.keyTap("tab");
    //await sleep(10);
  });
  robot.keyTap("end");
 robot.typeString('K');
};

main().catch((e) => {
  console.log(e);
  process.exit(1);
});
