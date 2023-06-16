const getExcelData = require("./getExcelData");
const sleep = require("./sleep");
const robot = require("./testRobot");

const range = (n) => Array.from(Array(n)).map((v, i) => i);

module.exports = {
  getExcelData,
  sleep,
  range,
  robot,
};
