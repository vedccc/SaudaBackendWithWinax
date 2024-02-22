const { connectionString } = require("../../config/database");
const moment = require("moment");
const winax = require("winax");

const ADODB = {
  CommandTypeEnum: {
    adCmdText: 1,
    adCmdStoredProc: 4,
  },
  ParameterDirectionEnum: {
    adParamInput: 1,
    adParamOutput: 2,
  },
  DataTypeEnum: {
    adInteger: 3,
    adDBTimeStamp: 135,
    adVarChar: 200,
  },
};

const adobeConnect = async (query) => {
  try {
    let connection = new winax.Object("ADODB.Connection");
    connection.Open(connectionString);
    let rs = new winax.Object("ADODB.Recordset");
    rs.Open(query, connection);
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

const Get_MaxExCondate = (exchangeId, toDate) => {
  toDate = formatDateYYYYMMDD(toDate);
  const lop = "0";

  let connection = new winax.Object("ADODB.Connection");
  connection.Open(connectionString);

  const command = new winax.Object("ADODB.Command");
  command.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc;
  command.CommandText = "Get_MaxExCondate";
  command.ActiveConnection = connection;

  command.Parameters.Append(
    command.CreateParameter(
      "@MC_CODE",
      ADODB.DataTypeEnum.adInteger,
      ADODB.ParameterDirectionEnum.adParamInput,
      -1,
      1005
    )
  );
  command.Parameters.Append(
    command.CreateParameter(
      "@EXID",
      ADODB.DataTypeEnum.adInteger,
      ADODB.ParameterDirectionEnum.adParamInput,
      -1,
      exchangeId
    )
  );
  command.Parameters.Append(
    command.CreateParameter(
      "@TODATE",
      ADODB.DataTypeEnum.adDBTimeStamp,
      ADODB.ParameterDirectionEnum.adParamInput,
      -1,
      toDate
    )
  );
  command.Parameters.Append(
    command.CreateParameter(
      "@LDATE",
      ADODB.DataTypeEnum.adVarChar,
      ADODB.ParameterDirectionEnum.adParamOutput,
      20,
      lop
    )
  );

  const rs = command.Execute();
  return rs;
};

const formatDateYYYYMMDD = (date) => {
  return moment(date).format("YYYY/MM/DD");
};

module.exports = {
  adobeConnect,
  formatDateYYYYMMDD,
  Get_MaxExCondate,
};
