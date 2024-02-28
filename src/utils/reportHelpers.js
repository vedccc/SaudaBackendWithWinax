const { connectionString } = require("../../config/database");
const { adobeConnect, formatDateYYYYMMDD } = require("./utils");
const winax = require("winax");

function calcBrokerage(
  LPBrokType,
  LPBrokRate,
  LPBrokRate2,
  LPBrokQty,
  LPQty,
  LPRATE,
  LPCalval,
  LPInstType,
  LPStrike,
  LPBrokQty2,
  LPTradeableLot,
  LPCconType,
  sno,
  lpitemid
) {
  let snoValue = sno || "1";
  let LPCBrokamt = 0;
  let GScGroup = ""; // Define GScGroup appropriately in your code

  const TRec2 = new winax.Object("ADODB.Recordset");
  const Cnn = new winax.Object("ADODB.Connection");
  Cnn.Open(connectionString);
  const mysql = `SELECT EXchangecode,QTYUNIT FROM ITEMMAST WHERE ITEMID = ${lpitemid}`;
  TRec2.Open(mysql, Cnn, 3, 1); // adOpenKeyset: 3, adLockReadOnly: 1

  if (TRec2.EOF) {
    return LPCBrokamt;
  }

  let EXCHANGECODE = TRec2.Fields("EXchangecode").Value;
  let QTYUNIT = TRec2.Fields("QTYUNIT").Value;

  if (LPBrokType !== "D") {
    if (LPBrokType === "T") {
      LPCBrokamt = LPBrokRate * LPQty;
    } else if (LPBrokType === "A") {
      if (LPBrokQty2 === 0) {
        console.log("Please Update Brokerage Qty");
      } else {
        LPCBrokamt = (LPBrokRate * LPBrokQty) / LPBrokQty2;
      }
    } else if (LPBrokType === "5") {
      if (LPBrokQty2 === 0) {
        // console.log("Please Update Brokerage Qty");
      } else {
        LPCBrokamt = (LPBrokRate * LPBrokQty) / LPBrokQty2;
      }
    } else if (LPBrokType === "1") {
      let A = (LPBrokRate / 100) * LPRATE;
      LPCBrokamt = A * LPQty;
    } else if (LPBrokType === "Z") {
      if (LPBrokQty === 0) {
        // console.log("Please Update Brokerage Qty");
      } else {
        if (EXCHANGECODE === "MCX" && QTYUNIT !== null && QTYUNIT >= 1) {
          LPCBrokamt = (LPBrokRate * LPQty) / QTYUNIT;
        } else {
          LPCBrokamt = LPBrokRate * (LPQty / LPBrokQty);
        }
      }
    } else if (LPBrokType === "R") {
      if (LPBrokQty2 === 0) {
        console.log("Please Update Brokerage Qty");
      } else {
        if (LPBrokQty !== 0) {
          LPCBrokamt = LPBrokRate * (LPBrokQty / LPBrokQty2);
        } else {
          LPCBrokamt = 0;
        }
      }
    } else if (LPBrokType === "F") {
      if (LPTradeableLot !== 0) {
        if (LPTradeableLot < LPQty) LPTradeableLot = LPQty;
        LPCBrokamt = LPBrokRate * (LPQty / LPTradeableLot);
      }
    } else if (LPBrokType === "P") {
      LPCBrokamt = (LPBrokRate / 100) * LPQty * LPRATE * LPCalval;
    } else if (LPBrokType === "2") {
      LPCBrokamt = (LPBrokRate / 100) * (LPQty * 100 * LPCalval);
    } else if (LPBrokType === "U") {
      LPCBrokamt = LPBrokRate * LPQty * LPTradeableLot;
    } else if (LPBrokType === "O" || LPBrokType === "C") {
      LPCBrokamt = LPBrokRate * LPBrokQty;
    } else if (LPBrokType === "N") {
      LPCBrokamt = LPBrokRate;
    } else if (LPBrokType === "I") {
      LPCBrokamt = (LPBrokRate / 100) * LPBrokQty * LPRATE * LPCalval;
      LPCBrokamt +=
        (LPBrokRate2 / 100) * (LPQty - LPBrokQty) * LPRATE * LPCalval;
    } else if (LPBrokType === "4") {
      LPCBrokamt = (LPBrokRate / 100) * LPBrokQty * LPRATE * LPCalval;
      LPCBrokamt += (LPBrokRate2 / 100) * LPBrokQty2 * LPRATE * LPCalval;
    } else if (LPBrokType === "Q") {
      LPCBrokamt = LPBrokRate * LPBrokQty;
    } else if (LPBrokType === "V") {
      LPCBrokamt = (LPBrokRate / 100) * LPBrokQty * LPRATE * LPCalval;
      LPCBrokamt +=
        (LPBrokRate2 / 100) * (LPQty - LPBrokQty) * LPRATE * LPCalval;
    } else if (LPBrokType === "S") {
      LPCBrokamt = LPBrokRate * LPBrokQty;
      LPCBrokamt += LPBrokRate2 * (LPQty - LPBrokQty);
    } else if (LPBrokType === "Y") {
      LPCBrokamt = LPBrokRate * LPBrokQty;
      LPCBrokamt += LPBrokRate2 * (LPQty - LPBrokQty);
    } else if (LPBrokType === "M") {
      if (LPBrokQty !== 0) {
        let A = LPRATE * LPBrokQty * LPCalval * (LPBrokRate / 100);
        let B = A / LPBrokQty;
        LPCBrokamt = Math.round(B * LPBrokQty * 100) / 100;
      }
      if (LPQty - LPBrokQty !== 0) {
        let A = (LPQty - LPBrokQty) * LPRATE * LPCalval * (LPBrokRate2 / 100);
        let B = A / (LPQty - LPBrokQty);
        LPCBrokamt += Math.round(B * (LPQty - LPBrokQty) * 100) / 100;
      }
    } else if (LPBrokType === "W") {
      LPCBrokamt = (LPBrokRate / 100) * (LPQty - LPBrokQty) * LPRATE * LPCalval;
      LPCBrokamt += (LPBrokRate2 / 100) * (LPBrokQty * LPRATE * LPCalval);
    } else if (LPBrokType === "X") {
      LPCBrokamt =
        Math.round(((LPRATE * LPBrokRate) / 100) * LPBrokQty * LPCalval * 100) /
        100;
    } else if (LPBrokType === "H") {
      LPCBrokamt = (LPBrokRate / 100) * LPBrokQty * LPRATE * LPCalval;
    } else if (LPBrokType === "L") {
      LPCBrokamt = LPBrokRate * LPBrokQty;
    }
  } else if (LPBrokType === "D") {
    if (GScGroup !== "T") {
      if (LPCconType === "S") {
        LPCBrokamt =
          (LPBrokRate / 100) * (LPQty - LPBrokQty) * LPRATE * LPCalval;
      }
      LPCBrokamt += (LPBrokRate2 / 100) * LPBrokQty * LPRATE * LPCalval;
    } else {
      LPCBrokamt = (LPBrokRate2 / 100) * LPQty * LPRATE * LPCalval;
    }
  }

  return LPCBrokamt;
}

function standingAmt(
  LPSaudaCode,
  LPExCode,
  LPParty,
  LPFromDate,
  LPToDate,
  LPMaturityDate,
  LPItemCode,
  LPSaudaID
) {
  let LStandingAmt = 0;
  let LMinDate;
  let LACCID;

  // Assuming Get_AccID and Get_MinCondate functions are defined elsewhere
  LACCID = Get_AccID(LPParty);
  let LStrMinDate = Get_MinCondate(LACCID, LPSaudaID, LPToDate);

  if (LStrMinDate.length > 0) {
    LMinDate = new Date(LStrMinDate);
  } else {
    return 0;
  }

  const Cnn = new winax.Object("ADODB.Connection");
  Cnn.Open(connectionString);
  const DateRec = new winax.Object("ADODB.Recordset");
  const RecCont = new winax.Object("ADODB.Recordset");

  let mysql = `SELECT * FROM SETTLE WHERE COMPCODE =${1005}`;
  mysql += `AND SETDATE >='${formatDateYYYYMMDD(LMinDate)}'`;
  mysql += `AND SETDATE >= ${formatDateYYYYMMDD(LPFromDate)}`;
  mysql += `AND SETDATE <='${formatDateYYYYMMDD(LPMaturityDate)}'`;
  mysql += `AND SETDATE <='formatDateYYYYMMDD(LPToDate)'`;
  mysql += `ORDER BY SETDATE`;

  DateRec.Open(mysql, Cnn, 0, 1); // adOpenForwardOnly: 0, adLockReadOnly: 1

  while (!DateRec.EOF) {
    let MStdQty = 0;
    let LSTDRate = 0;
    let MDt = DateRec.Fields("SETDATE").Value;

    MStdQty = SDOpQty(LPParty, LPSaudaID, new Date(MDt.getTime() + 86400000)); // Adding one day to MDt

    if (MDt <= LPMaturityDate) {
      if (MStdQty !== 0) {
        const TRec = new winax.Object("ADODB.Recordset");
        mysql = `SELECT STDRATE FROM PITBROK WHERE COMPCODE =${1005}`;
        mysql += `AND AC_CODE = '${LPParty}' AND INSTTYPE ='FUT'`;
        mysql += `AND ITEMCODE='${LPItemCode}'AND UPTOSTDT>='${formatDateYYYYMMDD(
          MDt
        )}'`;
        mysql += " ORDER BY UPTOSTDT";

        TRec.Open(mysql, Cnn, 0, 1);

        if (!TRec.EOF) {
          LSTDRate = parseFloat(TRec.Fields("STDRATE").Value);
        } else {
          mysql = `SELECT STDRATE FROM PEXBROK WHERE COMPCODE =${1005}`;
          mysql += `AND AC_CODE ='${LPParty}' AND INSTTYPE ='FUT'`;
          mysql += `AND EXCODE ='${LPExCode}' AND UPTOSTDT>='${formatDateYYYYMMDD(
            MDt
          )}'`;
          mysql += `ORDER BY UPTOSTDT`;

          TRec.Open(mysql, Cnn, 0, 1);

          if (!TRec.EOF) {
            LSTDRate = parseFloat(TRec.Fields("STDRATE").Value);
          }
        }

        LStandingAmt += Math.abs(MStdQty) * LSTDRate * -1;
      }
    }

    DateRec.MoveNext();
  }

  DateRec.Close();
  RecCont.Close();

  return LStandingAmt;
}

const Get_MaxExCondate = async (companyCode, exchangeId, toDateValue) => {
  try {
    const result = await adobeConnect(
      `DECLARE  @LDATE varchar(20);
       EXEC GET_MaxExCondate @MC_CODE =${companyCode}, @EXID =${exchangeId},@TODATE ='${formatDateYYYYMMDD(
        toDateValue
      )}',@LDATE = @LDATE OUTPUT;
       SELECT @LDATE AS LDATE`
    );
    return result.Fields("LDATE").Value;
  } catch (error) {
    console.error(error);
    throw error;
  }
};

const SDOpQty = async (AC_CODE, LTSaudaID, fromDateValue) => {
  try {
    const result = await adobeConnect(
      `DECLARE  @NETQTY FLOAT;
       EXEC SPOp_Qty @MC_CODE = 1003, @PARTY ='${AC_CODE}',@SAUDAID = ${LTSaudaID}, @UPTODATE ='${formatDateYYYYMMDD(
        fromDateValue
      )}',@NETQTY = @NETQTY OUTPUT;
       SELECT @NETQTY AS NETQTY;`
    );
    return result.Fields("NETQTY").Value;
  } catch (error) {
    console.error(error);
    throw error;
  }
};

const SDCLRATE = async (LTSaudaID, toDateValue, lotType) => {
  try {
    const result = await adobeConnect(
      `DECLARE  @LRATE FLOAT;
       EXEC GET_CLOSINGRATE @MC_CODE = 1005,@LCONDATE ='${formatDateYYYYMMDD(
         toDateValue
       )}',@SAUDAID = ${LTSaudaID},@LOTYPE = '${lotType}',@LRATE = @LRATE OUTPUT;
       SELECT @LRATE AS LRATE`
    );
    return result.Fields("LRATE").Value;
  } catch (error) {
    console.error(error);
    throw error;
  }
};

module.exports = {
  Get_MaxExCondate,
  SDOpQty,
  calcBrokerage,
  SDCLRATE,
  standingAmt,
};
