// Requiring the module
const XLSX = require("xlsx");
const fs = require("fs");

// Reading our file
const file = XLSX.readFile("./input.xlsx");

const sheets = file.SheetNames;

const data = XLSX.utils.sheet_to_json(file.Sheets[file.SheetNames[0]]);

let stringData = "";
data.map(async (item, index) => {
  console.log();
  stringData += `P${index + 1} ; ${item.X} ; ${item.Y}\n`;
});

fs.writeFile("tool-online-input.txt", stringData, () => {});

const toolOnlineResult = fs.readFileSync("tool-online-result.txt", "utf-8");

const resultArray = toolOnlineResult.split("\n");

resultArray.forEach((line, index) => {
  if (line?.length > 5) {
    const [a, latitude, longitude] = line.split(";");

    data[index].latitude = latitude - 1 + 1;
    data[index].longitude = longitude - 1 + 1;
  }
});

var worksheet = XLSX.utils.json_to_sheet(data);

const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet 1");
XLSX.writeFile(workbook, "output.xlsx");

console.log(data[0]);
