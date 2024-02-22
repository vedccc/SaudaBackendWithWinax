const { connectionString } = require("../../config/database");
const moment = require("moment");
var winax = require("winax");

const adobeConnect = async (query) => {
  try {
    let connection = new ActiveXObject("ADODB.Connection"); /*the line*/
    connection.Open(connectionString);
    let rs = new ActiveXObject("ADODB.Recordset");
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

module.exports = {
  adobeConnect,
  formatDateYYYYMMDD,
};
