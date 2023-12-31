const axios = require("axios");
const fs = require("fs");
const excel = require("excel4node");
require("dotenv").config();
process.env.TZ = "Europe/Zurich";

const exportInterventionHistory = async (firemanNIP) => {
	let currentData;
	let mirData;
	let fileModificationDate;

	// Get data from MIR
	try {
		const mirResponse = await axios.get(
			`https://${process.env.MIR_DOMAIN}/api/ecawin/statistics/fireman/${firemanNIP}`,
			{
				headers: {
					"Content-Type": "application/json",
					Authorization: `Bearer ${process.env.AUTH_TOKEN}`,
				},
			}
		);
		mirData = mirResponse.data;
	} catch (error) {
		console.error(error);
	}

	const fireman = mirData[0].firemen.filter((fireman) => fireman.nip === firemanNIP)[0];

	// Get data from file
	try {
		const dataFileLastModified = fs.statSync(
			`${process.env.OUTPUT_DIR}/${fireman.lastname}_${fireman.firstname}_${firemanNIP}_history.json`
		).mtimeMs;
		fileModificationDate = new Date(dataFileLastModified).toISOString().replace(/:/g, "-").split(".")[0];

		const currentDataFile = fs.readFileSync(
			`${process.env.OUTPUT_DIR}/${fireman.lastname}_${fireman.firstname}_${firemanNIP}_history.json`,
			"utf8"
		);
		currentData = JSON.parse(currentDataFile);
	} catch (error) {
		if (error.code === "ENOENT") {
			console.log("File not found, will be created");
			currentData = [];
		} else {
			console.error(error);
		}
	}

	// Append new data to existing JSON data
	console.log(currentData);
	console.log(mirData);

	for (const newIntervention of mirData.reverse()) {
		const isDuplicate = currentData.some((item) => {
			// Compare based on a unique identifier or specific fields
			return item.rapport === newIntervention.rapport;
		});

		if (isDuplicate) {
			console.log(
				`Intervention id ${newIntervention.rapport} (${newIntervention.what}) already exists. Skipping...`
			);
			continue;
		}

		// Prepend new data to existing JSON data
		console.log(`Adding intervention id ${newIntervention.rapport} (${newIntervention.what})`);
		currentData.unshift(newIntervention);
	}

	// Backup old file
	if (
		fs.existsSync(`${process.env.OUTPUT_DIR}/${fireman.lastname}_${fireman.firstname}_${firemanNIP}_history.json`)
	) {
		if (!fs.existsSync(`${process.env.OUTPUT_DIR}/.archives/`)) {
			fs.mkdirSync(`${process.env.OUTPUT_DIR}/.archives/`);
		}
		fs.renameSync(
			`${process.env.OUTPUT_DIR}/${fireman.lastname}_${fireman.firstname}_${firemanNIP}_history.json`,
			`${process.env.OUTPUT_DIR}/.archives/${fireman.lastname}_${fireman.firstname}_${firemanNIP}_${fileModificationDate}_history.json`
		);
	}

	// Export updated JSON data to file
	try {
		fs.writeFileSync(
			`${process.env.OUTPUT_DIR}/${fireman.lastname}_${fireman.firstname}_${firemanNIP}_history.json`,
			JSON.stringify(currentData),
			{
				encoding: "utf8",
				flag: "w",
			}
		);
	} catch (error) {
		console.error(error);
	}

	// Convert JSON to Excel
	const workbook = new excel.Workbook();
	const worksheet = workbook.addWorksheet("Historique");

	const headerStyle = workbook.createStyle({
		font: {
			bold: true,
		},
		alignment: {
			horizontal: "center",
		},
	});
	const dateStyle = workbook.createStyle({
		numberFormat: "dd.mm.yyyy hh:mm",
		alignment: {
			horizontal: "left",
		},
	});

	worksheet.cell(1, 1).string("Rapport").style(headerStyle);
	worksheet.cell(1, 2).string("Alarme").style(headerStyle);
	worksheet.cell(1, 3).string("Lieu").style(headerStyle);
	worksheet.cell(1, 4).string("Début intervention").style(headerStyle);
	worksheet.cell(1, 5).string("Fin intervention").style(headerStyle);
	worksheet.cell(1, 6).string("Durée").style(headerStyle);
	worksheet.cell(1, 7).string("CI").style(headerStyle);

	worksheet.column(1).setWidth(10);
	worksheet.column(2).setWidth(40);
	worksheet.column(3).setWidth(55);
	worksheet.column(4).setWidth(17);
	worksheet.column(5).setWidth(17);
	worksheet.column(6).setWidth(8);
	worksheet.column(7).setWidth(20);

	worksheet.row(1).freeze();

	for (let i = 0; i < currentData.length; i++) {
		const workRow = i + 2;
		const intervention = currentData[i];
		const interventionStartTime = new Date(intervention.alarmTime);
		const interventionStartTimeOffset = interventionStartTime.getTimezoneOffset();
		interventionStartTime.setMinutes(interventionStartTime.getMinutes() - interventionStartTimeOffset);
		const interventionEndTime = new Date(intervention.repliTime);
		const interventionEndTimeOffset = interventionEndTime.getTimezoneOffset();
		interventionEndTime.setMinutes(interventionEndTime.getMinutes() - interventionEndTimeOffset);
		const interventionDuration = new Date(interventionEndTime - interventionStartTime);
		const interventionDurationFormatted = `${interventionDuration.getUTCHours()}h${interventionDuration
			.getMinutes()
			.toString()
			.padStart(2, "0")}`;

		worksheet.cell(workRow, 1).string(intervention.rapport);
		worksheet.cell(workRow, 2).string(intervention.what);
		worksheet.cell(workRow, 3).string(`${intervention.where} - ${intervention.npa} ${intervention.localite}`);
		worksheet.cell(workRow, 4).date(interventionStartTime).style(dateStyle);
		worksheet.cell(workRow, 5).date(interventionEndTime).style(dateStyle);
		worksheet.cell(workRow, 6).string(interventionDurationFormatted);
		worksheet
			.cell(workRow, 7)
			.string(`${intervention.chief.rank} ${intervention.chief.firstname} ${intervention.chief.lastname}`);
	}

	// Backup old file
	if (
		fs.existsSync(`${process.env.OUTPUT_DIR}/${fireman.lastname}_${fireman.firstname}_${firemanNIP}_history.xlsx`)
	) {
		if (!fs.existsSync(`${process.env.OUTPUT_DIR}/.archives/`)) {
			fs.mkdirSync(`${process.env.OUTPUT_DIR}/.archives/`);
		}
		fs.renameSync(
			`${process.env.OUTPUT_DIR}/${fireman.lastname}_${fireman.firstname}_${firemanNIP}_history.xlsx`,
			`${process.env.OUTPUT_DIR}/.archives/${fireman.lastname}_${fireman.firstname}_${firemanNIP}_${fileModificationDate}_history.xlsx`
		);
	}

	// Export updated xslx
	try {
		workbook.write(`${process.env.OUTPUT_DIR}/${fireman.lastname}_${fireman.firstname}_${firemanNIP}_history.xlsx`);
	} catch (error) {
		console.error(error);
	}
};

(async () => {
	const firemen = process.env.FIREMEN_NIP.split(",");

	for (const fireman of firemen) {
		console.log(`Exporting data for ${fireman}`);
		await exportInterventionHistory(fireman);
		console.log("Done\n--------------------------------\n");
	}
})();
