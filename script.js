const xlsx = require("xlsx");
const path = require("path");
const wb = xlsx.readFile(path.join(__dirname + "/uploads/file-1575501862380.xlsx"), {
	type: "binary",
	cellDates: true,
	dateNF: "yyyy-mm-dd",
});
var ws1 = wb.Sheets["unit_station_names"];
var ws2 = wb.Sheets["loan_status"];
var ws3 = wb.Sheets["loans"];
const dataforws1 = xlsx.utils.sheet_to_json(ws1);
const dataforws2 = xlsx.utils.sheet_to_json(ws2);
const dataforws3 = xlsx.utils.sheet_to_json(ws3);
let values = [];
for (let i = 0; i < 10; i++) {
	values.push([dataforws1[i].station_name, dataforws1[i][" dailytarget "], dataforws1[i][" monthlytarget "]]);
}

const loansWB = xlsx.utils.book_new();
const newWS = xlsx.utils.json_to_sheet(dataforws3);
xlsx.utils.book_append_sheet(loansWB, newWS, "loans");
xlsx.writeFile(loansWB, "loans.xlsx");
