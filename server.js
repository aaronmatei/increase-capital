const express = require("express");
const path = require("path");
const xlsx = require("xlsx");
const app = express();
const bodyParser = require("body-parser");
const multer = require("multer");
const xlstojson = require("xls-to-json-lc");
const xlsxtojson = require("xlsx-to-json-lc");
const moment = require("moment");
const cors = require("cors");

const mysql = require("mysql");
const db = mysql.createConnection({
	host: "localhost",
	user: "root",
	password: "",
	database: "increase",
	dateStrings: true,
});
// use body-parser
app.use(cors());

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// set view engine to ejs

app.set("port", process.env.PORT || 5000);
app.engine("ejs", require("express-ejs-extend"));
app.set("view engine", "ejs");
app.set("views", path.join(__dirname + "/views"));

app.use(express.static(__dirname + "/public"));

const downloadPDF = require("./routes/downloads/downloadpdf");
const downloadXLS = require("./routes/downloads/downloadxls");
app.use(downloadPDF);
app.use(downloadXLS);

// Create DB
app.get("/createdb", (req, res) => {
	let sql = "CREATE DATABASE increase";
	db.query(sql, (err, result) => {
		if (err) throw err;
		res.send("created db successfully");
	});
});

// Home route
app.get("/", (req, res) => {
	res.render("home");
});

// Create table
app.get("/createtable", (req, res) => {
	let sql1 = "DROP TABLE IF EXISTS `unit_station_names`";
	db.query(sql1, (err, result) => {
		if (err) throw err;
		// res.send("Table Created Successfully");
	});
	let sql =
		"CREATE TABLE unit_station_names (id INT AUTO_INCREMENT PRIMARY KEY, station_name VARCHAR(255), daily_target INT, monthly_target INT)";
	db.query(sql, (err, result) => {
		if (err) throw err;
		res.send("Table Created Successfully");
	});
});
// Create loan status table
app.get("/createlstable", (req, res) => {
	let sql1 = "DROP TABLE IF EXISTS `loan_status`";
	db.query(sql1, (err, result) => {
		if (err) throw err;
		// res.send("Table Created Successfully");
	});
	let sql = "CREATE TABLE loan_status (id INT AUTO_INCREMENT PRIMARY KEY, loan_status VARCHAR(255))";
	db.query(sql, (err, result) => {
		if (err) throw err;
		res.send("Loan Status Table Created Successfully");
	});
});
// Create loan status table
app.get("/createloanstable", (req, res) => {
	let sql1 = "DROP TABLE IF EXISTS `loans`";
	db.query(sql1, (err, result) => {
		if (err) throw err;
		// res.send("Table Created Successfully");
	});
	let sql =
		"CREATE TABLE loans (loan_date DATETIME, due_date DATETIME, loan_code INT, loan_amount INT, customer_station INT, customer_id VARCHAR(255), loan_status INT, FOREIGN KEY (customer_station) REFERENCES unit_station_names (id),FOREIGN KEY (loan_status) REFERENCES loan_status (id))";
	db.query(sql, (err, result) => {
		if (err) throw err;
		res.send("Loans Table Created Successfully");
	});
});

// upload excel with data
const storage = multer.diskStorage({
	destination: (req, file, cb) => {
		cb(null, "./uploads/dbexcel");
	},
	filename: (req, file, cb) => {
		const datetimestamp = Date.now();
		cb(
			null,
			file.fieldname +
				"-" +
				datetimestamp +
				"." +
				file.originalname.split(".")[file.originalname.split(".").length - 1]
		);
	},
});

const upload = multer({
	storage: storage,
	fileFilter: (req, file, callback) => {
		// file filter
		if (["xls", "xlsx"].indexOf(file.originalname.split(".")[file.originalname.split(".").length - 1]) === -1) {
			return callback(new Error("Wrong extension type"));
		}
		callback(null, true);
	},
}).single("file");

// API that will upload the excel file
app.post("/upload", (req, res) => {
	let exceltojson;
	upload(req, res, err => {
		if (err) {
			res.json({ error_code: 1, err_desc: err });
			return;
		}
		// req.file contains the file info
		if (!req.file) {
			res.json({ error_code: 1, err_desc: "No file uploaded" });
		}
		// start the conversion process
		// check extension
		if (req.file.originalname.split(".")[req.file.originalname.split(".").length - 1] === "xlsx") {
			exceltojson = xlsxtojson;
		} else {
			exceltojson = xlstojson;
		}

		try {
			exceltojson(
				{
					input: req.file.path, //where the file was uploaded
					output: null, // we dont need output.json
					lowerCaseHeaders: true,
				},
				(err, result) => {
					if (err) {
						return res.json({ error_code: 1, err_desc: err, data: null });
					}
					res.json({ error_code: 0, err_desc: null, data: result });
					console.log(data);
				}
			);
		} catch (err) {
			res.json({ error_code: 1, err_desc: "Excel file uploaded is corrupted" });
		}

		res.json({ error_code: 0, err_desc: null });
	});
});

app.get("/uploadfile", (req, res) => {
	res.render("uploads");
});

// Insert Stations into the table
app.get("/insertsn", (req, res) => {
	const wb = xlsx.readFile(path.join(__dirname + "/uploads/dbexcel/file-1575875523773.xlsx"), { cellDates: true });
	var ws1 = wb.Sheets["unit_station_names"];

	const dataforws1 = xlsx.utils.sheet_to_json(ws1);

	let values1 = [];
	for (let i = 0; i < dataforws1.length; i++) {
		values1.push([dataforws1[i].station_name, dataforws1[i][" dailytarget "], dataforws1[i][" monthlytarget "]]);
	}

	let sql1 = "INSERT INTO unit_station_names (station_name,daily_target, monthly_target) VALUES ?";
	db.query(sql1, [values1], (err, result) => {
		if (err) throw err;
		res.send("Records inserted");
	});
});
// Insert loan status values into the table
app.get("/insertls", (req, res) => {
	const wb = xlsx.readFile(path.join(__dirname + "/uploads/dbexcel/file-1575875523773.xlsx"), { cellDates: true });

	var ws2 = wb.Sheets["loan_status"];

	const dataforws2 = xlsx.utils.sheet_to_json(ws2);

	let values2 = [];
	for (let i = 0; i < dataforws2.length; i++) {
		values2.push([dataforws2[i].loan_status]);
	}

	let sql2 = "INSERT INTO loan_status (loan_status) VALUES ?";
	db.query(sql2, [values2], (err, result) => {
		if (err) throw err;
		res.send("Records inserted");
	});
});
// Insert loans Values into the table
app.get("/insertloans", (req, res) => {
	const wb = xlsx.readFile(path.join(__dirname + "/uploads/dbexcel/file-1575875523773.xlsx"), { cellDates: true });

	var ws3 = wb.Sheets["loans"];

	const dataforws3 = xlsx.utils.sheet_to_json(ws3);

	let values3 = [];
	for (let i = 0; i < dataforws3.length; i++) {
		values3.push([
			dataforws3[i].loan_date,
			dataforws3[i].due_date,
			dataforws3[i].loan_code,
			dataforws3[i].loan_amount,
			dataforws3[i].customer_station,
			dataforws3[i].customer_id,
			dataforws3[i].loan_status,
		]);
	}

	let sql3 =
		"INSERT INTO loans (loan_date,due_date,loan_code,loan_amount,customer_station,customer_id,loan_status) VALUES ?";
	db.query(sql3, [values3], (err, result) => {
		if (err) throw err;
		res.send("Records inserted");
	});
});

// query station_names from the db
app.get("/stations", (req, res) => {
	let sql = "SELECT * FROM unit_station_names";
	db.query(sql, (err, results, fields) => {
		if (err) throw err;

		let name = "Aronique";
		res.render("stations", { name: name, data: results });
	});
});
// query loan status from the db
app.get("/loanstatus", (req, res) => {
	let sql = "SELECT * FROM loan_status";
	db.query(sql, (err, results, fields) => {
		if (err) throw err;

		let name = "Aronique";
		res.render("loan_status", { name: name, data: results });
	});
});
// query loans the db
app.get("/loans", (req, res) => {
	let sql = "SELECT * FROM loans";
	db.query(sql, (err, result, fields) => {
		if (err) throw err;

		let name = "Aronique";

		res.render("loans", { name: name, data: result, moment: moment });
	});
});

// Filtering
app.post("/filtered", (req, res) => {
	let from_date = req.body.from_date;
	let to_date = req.body.to_date;
	let customer_name = req.body.customer_name;
	let loan_date = req.body.loan_date;
	let due_date = req.body.due_date;
	let loan_code = req.body.loan_code;
	let loan_amount = req.body.loan_amount;
	let customer_station = req.body.customer_station;
	let customer_id = req.body.customer_id;
	let loan_status = req.body.loan_status;

	let sql = `SELECT * FROM loans WHERE DATE_FORMAT(loan_date, '%Y-%m-%d') = ? OR loan_amount = ? OR customer_station = ? && loan_status = ?`;

	db.query(sql, [loan_date, loan_amount, customer_station, loan_status], (err, result, fields) => {
		if (err) throw err;
		console.log(loan_date);

		res.render("results", { data: result, moment: moment });
	});
});

// Query by id
app.get("/results/:id", (req, res) => {
	let sql = `SELECT * FROM unit_station_names WHERE id=${req.params.id}`;
	db.query(sql, (err, result, fields) => {
		if (err) throw err;
		console.log(result);
	});
	res.send(`Record for id ${req.params.id}  fetched....`);
});
// Update record by id
app.get("/updateresult/:id", (req, res) => {
	let station_name = "Collections Cell";
	let sql = `UPDATE unit_station_names SET station_name = '${station_name}' WHERE id=${req.params.id}`;
	db.query(sql, (err, result, fields) => {
		if (err) throw err;
		console.log(result);
	});
	res.send(`Record for id ${req.params.id}  updated to ${station_name}....`);
});
// Delete record by id
app.get("/deleteresult/:id", (req, res) => {
	let sql = `DELETE * FROM unit_station_names WHERE id=${req.params.id}`;
	db.query(sql, (err, result, fields) => {
		if (err) throw err;
		console.log(result);
	});
	res.send(`Record for id ${req.params.id}  deleted....`);
});

app.listen(5000, () => {
	console.log("Server running on port 5000");
});
