const { connectionString } = require("../../config/database");
const moment = require("moment");
const winax = require("winax");

const adobeConnect = async (query) => {
  try {
    let connection = new winax.Object("ADODB.Connection");
    connection.Open(connectionString);
    let rs = new winax.Object("ADODB.Recordset");
    rs.Open(query, connection);
    return rs;
  } catch (error) {
    console.error(error);
    throw error;
  }
};

const formatDateYYYYMMDD = (date) => {
  return moment(date).format("YYYY/MM/DD");
};
const formatDateDDMMYYYY = (date) => {
  return moment(date).format("DD/MM/YYYY");
};

module.exports = {
  adobeConnect,
  formatDateYYYYMMDD,
  formatDateDDMMYYYY,
};
