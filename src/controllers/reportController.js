const partyModel = require("../models/partyModel");
const {
  generateExcel,
  generatePDF,
} = require("../utils/billSummaryExelHelper");
const { PassThrough } = require("stream");

const {
  SDCLRATE,
  SDOpQty,
  Get_MaxExCondate,
  calcBrokerage,
} = require("../utils/reportHelpers");
const { adobeConnect, formatDateYYYYMMDD } = require("../utils/utils");
const Excel = require("exceljs");

async function handleAllLogic(companyCode, AC_CODE, vcDTP2Value, fromDate) {
  let LDebitAmt = 0;
  let LCreditAmt = 0;
  let returnValue = 0;
  let mysql = `
        SELECT DR_CR, SUM(AMOUNT) AS AMT
        FROM VCHAMT
        WHERE COMPCODE = ${companyCode}
        AND AC_CODE = '${AC_CODE}'
        AND VOU_DT >= '${Format(fromDate, "yyyy/MM/dd")}'
        AND VOU_TYPE <> 'S'
        AND VOU_DT <= '${Format(vcDTP2Value, "yyyy/MM/dd")}'
        GROUP BY DR_CR
    `;

  TRec = null;
  TRec = await adobeConnect(mysql);
  if (!TRec.EOF) {
    if (TRec.DR_CR === "C") LCreditAmt = IsNull(TRec.AMT, 0);
    if (TRec.DR_CR === "D") LDebitAmt = IsNull(TRec.AMT, 0);
    TRec.MoveNext();
    if (!TRec.EOF) {
      if (TRec.DR_CR === "C") LCreditAmt = IsNull(TRec.AMT, 0);
      if (TRec.DR_CR === "D") LDebitAmt = IsNull(TRec.AMT, 0);
    }
    returnValue = LCreditAmt - LDebitAmt;
  }
  TRec = null;
  return returnValue;
}

const reportController = {
  async billSummary(req, res) {
    try {
      let LprevBal,
        LCurrFRate,
        LBal,
        LMinRate,
        LMStrDate,
        temporaryRecordSet,
        TRec,
        LBrokRate,
        LBrokRate2,
        LCalval,
        LCloseRate,
        LPartyName,
        LOpenRate,
        LTotBuyQty,
        LTotSellQty,
        LTotBuyAmt,
        LTotSellAmt,
        LTotOpAmt,
        lOpQty,
        LBrokType,
        CountRec,
        LTillDate,
        LTFromDate,
        LSetNo,
        LBillBy,
        LNetRecdPay,
        LNetMargin,
        AccStAdoREC,
        LStampDutyDate,
        LSaudaCode_muliple,
        mpartyop,
        mpusdinr;

      const partiesRecordset = await partyModel.getPartyRecordsBySelection(req);

      let arr = [];
      // Party Loop Start //
      while (!partiesRecordset.EOF) {
        let {
          AC_CODE,
          NAME,
          ACCID,
          OP_BAL,
          SRTAXAPP,
          RISKMTYPE,
          SEBITYPE,
          CGST,
          SGST,
          IGST,
          UTT,
          APPLYON,
        } = partiesRecordset.Fields;
        const tradableItemRecordSet =
          await partyModel.getTradedItemDetailsBySelection(req);

        let netPosition = 0;
        let brokAmount = 0;
        LPartyName = String(NAME);

        if (req?.body?.GOnlyBrok === 0) {
          SRTAXAPP = SRTAXAPP;
          RISKMTYPE = RISKMTYPE === null ? "N" : RISKMTYPEe;
          SEBITYPE = SEBITYPE === null ? "N" : SEBITYPE;
          CGST = CGST === null ? "N" : CGST;
          LSGSTType = SGST === null ? "N" : SGST;
          LIGSTType = IGST === null ? "N" : IGST;
          LUTTType = UTT === null ? "N" : UTT;
          LPStmType = APPLYON === null ? "N" : APPLYON;
        }
        LTillDate = new Date(req?.body?.toDate);
        LTFromDate = new Date(req?.body?.fromDate);
        LMStrDate = "";
        // LMStrDate = Get_MaxExCondate(1005, EXID, LTillDate);
        // let Lmaturity = MATURITY;
        LMStrDate.length > 1 && (LTillDate = new Date(LMStrDate));
        // LSaudaCode_muliple =
        //   LSaudaCode_muliple === ""
        //     ? SAUDACODE
        //     : `${LSaudaCode_muliple};${SAUDACODE}`;
        // REFLOT = REFLOT != null ? REFLOT : 1;
        // BROKLOT = BROKLOT != null ? BROKLOT : 1;
        LMinRate = 0;
        if (req?.body?.GMinBrokYN === "Y" && req?.body?.GFBroktype === "Y") {
          // getminrate(AC_CODE, ITEMID, EXID, LMinRate);
        }
        LBuyClRate = 1;
        LSellClRate = 1;

        if (req?.body?.LOpFlag === false) {
          LNetRecdPay = 0;

          if (req?.body?.confirmed || req?.body?.all) {
            LprevBal = OP_BAL;
            if (req?.body?.Check10 === 0) {
              LBal = Net_DrCr(AC_CODE, req?.body?.fromDate);
              LprevBal += LBal;
            }
            if (req?.body?.confirmed) {
              if (req?.body?.showTradeConfirm) {
                LNetRecdPay = 0;
                LNetMargin = 0;
                TRec = null;
                let mysql = `SELECT SUM(CASE DR_CR  WHEN 'D' THEN AMOUNT * -1  WHEN 'C' THEN AMOUNT * 1 END) AS AMT
                             FROM VCHAMT
                             WHERE COMPCODE = ${body?.rec?.companyCode}  AND AC_CODE = '${partiesRecordset?.AC_CODE}' AND VOU_DT < '${body?.rec?.fromDate}' AND CHEQUE_NO NOT IN ('DMARGIN', 'RMARGIN');`;

                TRec = await adobeConnect(mysql);
                if (!TRec.EOF) LprevBal += IsNull(TRec.AMT, 0);

                mysql = `SELECT SUM(CASE DR_CR WHEN 'D' THEN AMOUNT * -1 WHEN 'C' THEN AMOUNT * 1 END) AS AMT FROM VCHAMT WHERE COMPCODE = ${
                  req?.body?.companyCode
                } AND AC_CODE = '${AC_CODE}' AND VOU_DT >= '${Format(
                  req?.body?.fromDate,
                  "yyyy/MM/dd"
                )}' AND VOU_DT <= '${Format(
                  vcDTP2.Value,
                  "yyyy/MM/dd"
                )}' AND VOU_TYPE <> 'S' AND CHEQUE_NO NOT IN ('DMARGIN', 'RMARGIN')
                       `;

                TRec = await adobeConnect(mysql);
                if (!TRec.EOF) LNetRecdPay = IsNull(TRec.AMT, 0);

                mysql = `SELECT IMARGIN FROM DLYCLMGN WHERE COMPCODE = ${
                  req?.body?.companyCode
                } AND CLIENT = '${AC_CODE}' AND CONDATE = '${Format(
                  vcDTP2.Value,
                  "yyyy/MM/dd"
                )}'AND EXCODE = 'CME'
                `;
                TRec = await adobeConnect(mysql);
                if (!TRec.EOF) LNetMargin = IsNull(TRec.IMARGIN, 0);

                TRec = null;
              } else {
                if (req?.body?.withSharing) {
                  LprevBal += Get_Balance(ACCID);
                } else {
                  TRec = null;
                  let mysql = `EXEC ACSUMVCHS ${req?.body?.companyCode}, '${AC_CODE}', '${req?.body?.fromDate}', '${req?.body?.toDate}'`;
                  TRec = await adobeConnect(mysql);
                  if (!TRec.EOF) LprevBal += IsNull(TRec.AMT, 0);
                  TRec = null;
                }
              }
            } else if (req?.body?.all) {
              LNetRecdPay = handleAllLogic(
                req?.body?.companyCode,
                AC_CODE,
                req?.body?.toDate,
                req?.body?.fromDate
              );
            }
          }
          LOpFlag = true;
        }

        while (!tradableItemRecordSet.EOF) {
          let {
            EXID,
            EXCODE,
            INSTTYPE,
            SAUDACODE,
            MATURITY,
            REFLOT,
            BROKLOT,
            LOTWISE,
            SAUDAID,
            ITEMID,
            TRADEABLELOT,
            LOT,
            STRIKEPRICE,
          } = tradableItemRecordSet.Fields;

          LCalval =
            (EXCODE === "NSE" && LOTWISE === "Y") ||
            (EXCODE === "MCX" && LOTWISE === "Y")
              ? TRADEABLELOT
              : LOT;

          mpusdinr = 1;

          temporaryRecordSet = null;
          temporaryRecordSet = await adobeConnect(
            `SELECT BILLBY,ACEXCODE FROM ACCT_EX WHERE COMPCODE = ${req?.body?.companyCode} AND AC_CODE = '${AC_CODE}' AND EXCODE = '${EXCODE}'`
          );
          if (!temporaryRecordSet.EOF) {
            LBillBy = temporaryRecordSet.Fields("BILLBY").Value;
          }
          temporaryRecordSet = null;

          if (EXCODE === "CMX") {
            let TRec = await adobeConnect(
              `SELECT closerate, HGRATE, LOWRATE FROM CTR_R WHERE COMPCODE = ${
                req?.body?.companyCode
              } AND SAUDA = 'USDINR' AND CONDATE = '${formatDateYYYYMMDD(
                req?.body?.toDate
              )}'`
            );
            if (!TRec.EOF) {
              LCurrFRate = TRec.CLOSERATE;
              LCurrCRate = TRec.LOWRATE;
              LCurrDRate = TRec.HGRATE;
            }
          } else {
            LCurrDRate = LCurrCRate = 1;
          }

          switch (LBillBy) {
            case "B":
              let TRec = null;
              TRec = await adobeConnect(
                `SELECT CLOSERATE, HGRATE, LOWRATE FROM CTR_R WHERE COMPCODE = ${
                  req?.body?.companyCode
                } AND SAUDAID = '${SAUDAID}' AND CONDATE = '${formatDateYYYYMMDD(
                  req?.body?.toDate
                )}'`
              );
              if (!TRec.EOF) {
                LBuyClRate = TRec.LOWRATE === 0 ? TRec.CLOSERATE : TRec.LOWRATE;
                LSellClRate = TRec.HGRATE === 0 ? TRec.CLOSERATE : TRec.HGRATE;
              }
              break;
            case "S":
              LCurrCRate = LCurrFRate;
              LCurrDRate = LCurrFRate;
              break;
            case "P":
              let LPOPRATE = 0;
              let LPCLRATE = 0;
              TRec = null;
              TRec = await adobeConnect(
                `SELECT TOP 1 SETTLERATE FROM CTR_RP WHERE party = '${AC_CODE}' AND SAUDAID = '${SAUDAID}' AND CONDATE < '${formatDateYYYYMMDD(
                  req?.body?.fromDate
                )}' ORDER BY CONDATE DESC`
              );
              if (!TRec.EOF) {
                if (TRec.SETTLERATE !== 0) LPOPRATE = TRec.SETTLERATE;
              }

              TRec = await adobeConnect(
                `SELECT SETTLERATE FROM CTR_RP WHERE PARTY = '${AC_CODE}' AND SAUDAID = '${SAUDAID}' AND CONDATE = '${Format(
                  vcDTP2.Value,
                  "yyyy/mm/dd"
                )}'`
              );
              if (!TRec.EOF) {
                if (TRec.SETTLERATE !== 0) LPCLRATE = TRec.SETTLERATE;
              }

              TRec = await adobeConnect(
                `SELECT SETTLERATE, HiGh, LOW FROM CTR_RP WHERE COMPCODE = ${
                  req?.body?.companyCode
                } AND PARTY = '${AC_CODE}' AND SAUDA = 'USDINR' AND CONDATE = '${Format(
                  vcDTP2.Value,
                  "YYYY/MM/DD"
                )}'`
              );
              if (!TRec.EOF) {
                LCurrFRate = TRec.SETTLERATE;
                LCurrCRate = TRec.LOW;
                LCurrDRate = TRec.HIGH;
              }
              break;
          }

          // chek check 13 //
          if (req?.body?.GUniqClientId === "BRO2-CHE") {
            if (
              mpusdinr === 1 &&
              EXCODE === "CMX" &&
              Check5.Value === 0 &&
              Check13.Value === 1
            ) {
              let TRec = null;
              TRec = await adobeConnect(
                `SELECT CURRRATE, USDINR FROM INV_D WHERE COMPCODE = ${
                  req?.body?.companyCode
                } AND PARTY = '${AC_CODE}' AND EXCODE = 'CMX' AND STDATE >= '${Format(
                  req?.body?.fromDate,
                  "YYYY/MM/DD"
                )}' ORDER BY STDATE DESC`
              );
              if (!TRec.EOF) {
                mpusdinr = TRec.USDINR;
                LCurrDRate = TRec.CURRRATE;
              }
            }
            if (mpusdinr !== 1 || (Check5.Value === 0 && Check13.Value === 0)) {
              LCurrFRate = mpusdinr;
              LCurrCRate = mpusdinr;
            }
          } else {
            if (
              mpusdinr === 1 &&
              EXCODE === "CMX" &&
              Check5.Value === 0 &&
              Check13.Value === 1
            ) {
              let TRec = null;
              TRec = await adobeConnect(
                `SELECT CURRRATE, USDINR FROM INV_D WHERE COMPCODE = ${
                  req?.body?.companyCode
                } AND PARTY = '${AC_CODE}' AND EXCODE = 'CMX' AND STDATE >= '${Format(
                  req?.body?.fromDate,
                  "YYYY/MM/DD"
                )}' ORDER BY STDATE DESC`
              );
              if (!TRec.EOF) {
                mpusdinr = TRec.USDINR;
                LCurrDRate = TRec.USDINR;
              }
              if (
                mpusdinr !== 1 ||
                (Check5.Value === 0 && Check13.Value === 0)
              ) {
                LCurrFRate = mpusdinr;
                LCurrCRate = mpusdinr;
              }
            }
          }
          if (EXCODE === "CMX") {
            // LCurrCRate = 1;
            // LCurrDRate = 1;
          }
          LRiskMAmt = 0;
          LSBCTaxAmt = 0;
          LBrokAmount = 0;
          GSrvTax = 0;
          LTranAmount = 0;
          LStmAmt = 0;
          LTotStmAmt = 0;
          LStmRate = 0;
          LSTTAmt = 0;
          LTotSTTAmt = 0;
          LSTTRate = 0;
          LSEBITaxAmt = 0;
          LTotSEBITaxAmt = 0;
          LSEBITax = 0;
          LMarAmt = 0;
          LNStmAmt = 0;
          LCloseRate = 0;
          LClQty = 0;
          LNetPosition = 0;
          LEQSTTRate = 0;
          LEQSTMRate = 0;

          // check AcSt method //
          AccStAdoREC = null;
          AccStAdoREC = await adobeConnect(
            `EXEC AcSt ${req?.body?.companyCode},'${formatDateYYYYMMDD(
              LTFromDate
            )}','${formatDateYYYYMMDD(LTillDate)}',${SAUDAID},'${AC_CODE}' ,${
              req?.body?.GOnlyBrok
            }`
          );

          if (req?.body?.GOnlyBrok == 0) {
            if (APPLYON == "R" || APPLYON == "P") {
              if (LTillDate >= GSTMDate && LTillDate < LStampDutyDate) {
                LNStmAmt = Get_StampDutyAmt(
                  AC_CODE,
                  LTFromDate,
                  LTillDate,
                  LSExCodes
                );
              }
            }
          }

          LTotBuyQty = 0;
          LTotSellQty = 0;
          GSTDChrs = 0;
          lOpQty = 0;
          LBrokRate2 = 0;
          LCGSTAmt = 0;
          LSGSTAmt = 0;
          LIGSTAmt = 0;
          LUTTAmt = 0;
          LTotBuyAmt = 0;
          LTotSellAmt = 0;
          LTotOpAmt = 0;
          MSrvTax = 0;
          LATranType = "P";
          LBrokType = "P";
          LOpenRate = 0;
          LTotBQAmt = 0;
          LTotSQAmt = 0;
          LDiffAmt = 0;

          lOpQty = await SDOpQty(AC_CODE, SAUDAID, LTFromDate);

          lOpQty = Math.round(lOpQty * 100) / 100; // Round to 2 decimal places
          LNetOpenPosition = lOpQty;
          LNetPosition = lOpQty;
          LBalQty = lOpQty;
          if (lOpQty !== 0) {
            LOpenRate = await SDCLRATE(SAUDAID, LTFromDate - 1, "O"); // closing rate for opening
            // if (LPOPRATE !== 0) {
            //   LOpenRate = LPOPRATE;
            // }
            LTotOpAmt = lOpQty * LOpenRate * LCalval;
          }

          if (
            (INSTTYPE === "OPT" && OptMTMChk.Value === 0) ||
            (INSTTYPE === "CSH" && CshMTMchk.Value === 0)
          ) {
            lOpQty = 0;
            LTotOpAmt = 0;
          }

          TRec = null;
          TRec = new ActiveXObject("ADODB.Recordset");
          TRec = AccStAdoREC;

          while (!TRec.EOF) {
            let {
              BROKRATE2,
              BROKQTY,
              BROKRATE,
              BROKQTY2,
              ORDNO,
              QTY,
              RATE,
              CONTYPE,
              CONDATE,
              BROKTYPE,
              CONTIME,
              SRVTAX,
              TRANTYPE,
              TRANRATE,
              SBC_TAX,
              CGSTRATE,
              SGSTRATE,
              IGSTRATE,
              UTTRATE,
              SEBITAX,
              STTRATE,
              EQ_STT,
              EQ_STAMP,
              STMRATE,
              PATTAN,
            } = TRec.Fields;

            // b
            LCBrokAmt = 0;
            LcTranAmt = 0;

            LBrokType = BROKTYPE ? BROKTYPE : "T";
            LBrokRate2 = parseFloat(BROKRATE2) || 0;
            LBrokQty = BROKQTY != null ? BROKQTY : 0;
            LBrokRate = parseFloat(BROKRATE) || 0;
            LBrokQty2 = BROKQTY2 != null ? BROKQTY2 : 0;
            if (ORDNO !== null && ORDNO === "Carry") {
              LBrokRate = LBrokRate2;
            }

            if (CONTYPE == "B") {
              lOpQty += TRec.Fields("QTY").Value;
              LTotOpAmt += QTY * RATE * LCalval;
              switch (LBrokType) {
                case "Z":
                  LBrokQty = LBrokLot;
                  break;
                case "A":
                  LBrokQty2 = LBrokLot;
                  break;
                case "R":
                  LBrokQty2 = LBrokLot;
                  break;
                case "5":
                  LBrokQty2 = LBrokLot;
                  break;
                default:
                  break;
              }

              LCBrokAmt = calcBrokerage(
                LBrokType,
                LBrokRate,
                LBrokRate2,
                LBrokQty,
                QTY,
                RATE,
                LCalval,
                INSTTYPE,
                STRIKEPRICE,
                LBrokQty2,
                TRADEABLELOT,
                CONTYPE,
                "2",
                ITEMID
              );
              brokAmount = brokAmount + LCBrokAmt;
            } else if (CONTYPE == "S") {
              lOpQty -= TRec.Fields("QTY").Value;
              LTotOpAmt -= QTY * RATE * LCalval;

              switch (LBrokType) {
                case "Z":
                  LBrokQty = LBrokLot;
                  break;
                case "R":
                case "A":
                case "5":
                  LBrokQty2 = LBrokLot;
                  break;
                default:
                  // Handle default case if needed
                  break;
              }
              if (
                LBrokType == "O" ||
                LBrokType == "C" ||
                LBrokType == "A" ||
                LBrokType == "5"
              ) {
                // LBrokQty = Get_BrokQty(LBrokType, LBalQty, "S", QTY);
              }
              LCBrokAmt = calcBrokerage(
                LBrokType,
                LBrokRate,
                LBrokRate2,
                LBrokQty,
                QTY,
                RATE,
                LCalval,
                INSTTYPE,
                STRIKEPRICE,
                LBrokQty2,
                TRADEABLELOT,
                CONTYPE,
                "2",
                ITEMID
              );
              brokAmount = brokAmount + LCBrokAmt;
            }
            TRec.MoveNext();
          }
          LCloseRate = await SDCLRATE(SAUDAID, LTillDate, "C");
          let netamount = lOpQty * LCloseRate * LCalval;
          if (lOpQty > 0 && LTotOpAmt > 0) {
            LTotBuyQty = lOpQty;
            LTotBuyAmt = LTotOpAmt;
          } else {
            LTotSellQty = Math.abs(lOpQty);
            LTotSellAmt = Math.abs(LTotOpAmt);
          }
          if (LTotBuyQty > LTotSellQty && LTotBuyAmt > LTotSellAmt) {
            LTotSellAmt = LTotSellAmt + Math.abs(netamount);
          } else {
            LTotBuyAmt = LTotBuyAmt + Math.abs(netamount);
          }
          let Profit = 0;
          let Loss = 0;
          if (LTotBuyAmt > LTotSellAmt) {
            Loss = LTotBuyAmt - LTotSellAmt;
          } else {
            Profit = LTotSellAmt - LTotBuyAmt;
          }

          if (Profit > 0) {
            netPosition = netPosition + Profit;
          } else if (Loss > 0) {
            netPosition = netPosition - Loss;
          }

          tradableItemRecordSet.MoveNext();
        }

        if (netPosition > 0) {
          arr.push({
            partyName: LPartyName,
            gross_MTM: netPosition,
            brokAmt: brokAmount,
            netCreditAmount: netPosition - brokAmount,
          });
        } else {
          arr.push({
            partyName: LPartyName,
            gross_MTM: Math.abs(netPosition),
            brokAmt: brokAmount,
            netDebitAmount: Math.abs(netPosition) + brokAmount,
          });
        }

        partiesRecordset.MoveNext();
      }

      generatePDF(arr);
      const excelBuffer = await generateExcel(arr, LTFromDate, LTillDate);
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader("Content-Disposition", "attachment; filename=output.xlsx");

      res.send(Buffer.from(excelBuffer));
    } catch (error) {
      console.error("Error:", error);
      res.status(500).json({ message: "Internal Server Error" });
    }
  },
};

module.exports = reportController;
