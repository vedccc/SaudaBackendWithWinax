const express = require("express");
const router = express.Router();
// const partyController = require("../controllers/partyController");
// const companyController = require("../controllers/companyController");
const testController = require("../controllers/testControler");

// // parties
// router.get("/getPartiesBySelection", partyController.getPartiesBySelection);
// router.get("/getPartyRecordsBySelection", partyController.getPartyRecordsBySelection);

// // company
// router.get("/getAllCompanies", companyController.getAllCompanies);

// test
router.get("/test", testController.testStatment);

module.exports = router;
