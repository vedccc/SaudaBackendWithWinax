const express = require("express");
const reportController = require("../controllers/reportController");
const router = express.Router();
// const partyController = require("../controllers/partyController");
// const companyController = require("../controllers/companyController");

// // parties
// router.get("/getPartiesBySelection", partyController.getPartiesBySelection);
// router.get("/getPartyRecordsBySelection", partyController.getPartyRecordsBySelection);

// // company
// router.get("/getAllCompanies", companyController.getAllCompanies);

// test
router.post("/billSumary", reportController.billSummary);

module.exports = router;
