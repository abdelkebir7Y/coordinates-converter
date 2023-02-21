// Requiring the module
const XLSX = require("xlsx");
const fs = require("fs");
const proj4 = require("proj4");

const NATURE_VARIANTS = ["eau de surface", "Eau souterraine", "Lac et barrage"];

const DATA_VARIANTS_INITIAL_ACC = NATURE_VARIANTS.reduce(
  (acc, nature) => ({ ...acc, [nature]: [] }),
  { others: [] }
);

// Reading our input file
const file = XLSX.readFile("./input.xlsx");
const data = XLSX.utils.sheet_to_json(file.Sheets[file.SheetNames[0]]);

//write tool-online-input.txt files
let stringData = "";
data.forEach((item) => {
  stringData += `ID${item.ID} ; ${item.X} ; ${item.Y}\n`;
});
fs.writeFile("tool-online-input.txt", stringData, () => {});

// Reading our tool-online-result.txt file
const toolOnlineResult = fs.readFileSync("tool-online-result.txt", "utf-8");
const resultArray = toolOnlineResult.split("\n");
resultArray.forEach((line, index) => {
  if (line?.length > 5) {
    const [a, latitude, longitude] = line.split(";");
    data[index].latitude = latitude - 1 + 1;
    data[index].longitude = longitude - 1 + 1;
  }
});

//write output files

const dataVariants = data.reduce((acc, item) => {
  if (acc?.[item?.Nature ?? "d"]) {
    acc[item.Nature] = [...acc[item.Nature], item];
  } else {
    acc.others = [...acc.others, item];
  }

  return acc;
}, DATA_VARIANTS_INITIAL_ACC);

const workbook = XLSX.utils.book_new();

Object.entries(dataVariants).forEach(([prop, value]) => {
  if (value.length) {
    var worksheet = XLSX.utils.json_to_sheet(value);
    XLSX.utils.book_append_sheet(workbook, worksheet, prop);
  }
});

XLSX.writeFile(workbook, "output.xlsx");

console.log(data[0]);
