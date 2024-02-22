const sql = require("mssql/msnodesqlv8");
const ADODB = require("node-adodb");
const { connectionString } = require("../../config/database");
const moment = require("moment");
const winax = require("winax");
const connection2 = ADODB.open(connectionString);

const adobeConnect = async (query) => {
  const adOpenStatic = 3; // Cursor type
  const adLockReadOnly = 1;
  try {
    let connection = new winax.Object("ADODB.Connection");
    connection.Open(connectionString);
    let rs = new winax.Object("ADODB.Recordset");
    rs.Open(query, connection, 0, 1);
    if (!rs.EOF) {
      console.log("Recordset is full");
    } else {
      console.log("Recordset is empty.");
    }
    return rs;
  } catch (error) {
    console.error(error);
    throw error;
  }
};

const adobeConnect2 = async (query) => {
  try {
    const result = await connection2.query(query);

    if (result?.length > 0) {
      if (result.length > 1) {
        return result;
      } else {
        return result;
      }
    } else {
      return null;
    }
  } catch (error) {
    console.error(error);
    throw error;
  }
};

const formatDateYYYYMMDD = (date) => {
  return moment(date).format("YYYY/MM/DD");
};

const Get_MaxExCondate = async (companyCode, exchangeId, toDateValue) => {
  try {
    const result = await adobeConnect2(
      `DECLARE  @LDATE varchar(20);
       EXEC GET_MaxExCondate @MC_CODE = ${companyCode}, @EXID = '${exchangeId}', @TODATE ='${formatDateYYYYMMDD(
        toDateValue
      )}',@LDATE = @LDATE OUTPUT;
       SELECT @LDATE AS LDATE;`
    );
    return result[0]?.LDATE;
  } catch (error) {
    console.error(error);
    throw error;
  }
};

const SDOpQty = async (AC_CODE, LTSaudaID, toDateValue) => {
  try {
    const result = await adobeConnect(
      `DECLARE  @NETQTY FLOAT;
       EXEC SPOp_Qty @MC_CODE = 1003, @PARTY = '${AC_CODE}',@SAUDAID = ${LTSaudaID}, @UPTODATE ='${formatDateYYYYMMDD(
        toDateValue
      )}',@NETQTY = @NETQTY OUTPUT;
       SELECT @NETQTY AS NETQTY;`
    );
    return result[0]?.NETQTY;
  } catch (error) {
    console.error(error);
    throw error;
  }
};

const SDCLRATE = async (LTSaudaID, toDateValue, lotType) => {
  try {
    const result = await adobeConnect(
      `DECLARE  @LRATE FLOAT;
       EXEC GET_CLOSINGRATE @MC_CODE = 1003,@LCONDATE ='${formatDateYYYYMMDD(
         toDateValue
       )}',@SAUDAID = ${LTSaudaID}, @LOTYPE = '${lotType}',@LRATE = @LRATE OUTPUT;
       SELECT @LRATE AS LRATE;`
    );
    return result[0]?.LRATE;
  } catch (error) {
    console.error(error);
    throw error;
  }
};

module.exports = {
  adobeConnect,
  formatDateYYYYMMDD,
  Get_MaxExCondate,
  SDOpQty,
  SDCLRATE,
};
