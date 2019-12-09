const express = require("express");
const puppeteer = require("puppeteer");
const router = express.Router();

router.post("/pdfdownload", (req, res) => {
	res.send("Hey");
});
router.get("/pdfdownload", (req, res) => {
	const createPdf = async () => {
		const browser = await puppeteer.launch();
		const page = await browser.newPage();
		const options = {
			path: "uploads/pdfs/results.pdf",
			format: "A4",
		};
		await page.goto("http://localhost:5000/loans", { waitUntil: "networkidle2" });
		await page.pdf(options);
		await browser.close();
	};
	createPdf();
	res.send(`pdf created successfully. Check ~/uploads/pdfs/results.pdf `);
});

module.exports = router;
