const express = require("express");
const fs = require("fs");
const nodexlsx = require("node-xlsx");
const router = express.Router();
const mysql = require("mysql");
const db = mysql.createConnection({
	host: "localhost",
	user: "root",
	password: "",
	database: "increase",
	dateStrings: true,
});

router.get("/xlsdownload", (req, res) => {
	let sql = "SELECT * FROM loans LIMIT 50";
	db.query(sql, (err, result, fields) => {
		if (err) throw err;

		const data = json2Array(result, fields);
		const buffer = nodexlsx.build([{ name: "Loans", data: data }]);
		// Write the buffer to a file
		fs.writeFile("uploads/xls/loansdata.xlsx", buffer, fs_err => {
			if (fs_err) throw fs_err;
		});
	});
	res.send("Excel File created");
});

// function to convert json to array

const json2Array = (result, fields) => {
	let out = [];
	let temp = [];
	// Create headers array
	fields.forEach(item => {
		temp.push(item.name);
	});
	// temp array works as column headers in .xlsx file
	out.push(temp);

	result.forEach(item => {
		out.push([
			item.loan_date,
			item.due_date,
			item.loan_code,
			item.loan_amount,
			item.customer_station,
			item.customer_id,
			item.loan_status,
		]);
	});

	return out;
};

module.exports = router;
