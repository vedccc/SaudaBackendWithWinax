const {
  Net_DrCr,
  adobeConnect,
  adobeExcute,
  SDOpQty,
  getClosingRate,
  formatDateYYYYMMDD,
  SDCLRATE,
  Get_MaxExCondate,
  adobeConnect3,
  calcBrokerage,
  standingAmt,
} = require("../utils/utils");

const testController = {
  async testStatment(req, res) {
    try {
      let LprevBal,
        LPOPRATE,
        LCurrFRate,
        LBal,
        LDebitAmt,
        LCreditAmt,
        LMinRate,
        LMStrDate,
        TRec2,
        LPartyCode,
        LSaudaID,
        LExID,
        LItemID,
        LACCID,
        TRec,
        LStrike,
        LRefLot,
        LBuyAmt,
        LSellAmt,
        LBrokLot,
        LBrokRate,
        LBrokRate2,
        LBrokQty,
        LUTTAmt,
        LBrokQty2,
        LCalval,
        LTradeableLot,
        LMarRate,
        LCloseRate,
        LOpenRate,
        GSrvTax,
        LTranAmount,
        LStmAmt,
        LTotStmAmt,
        LStmRate,
        LSTTAmt,
        LTotSTTAmt,
        LSTTRate,
        LRiskMAmt,
        LSEBITaxAmt,
        LTotSEBITaxAmt,
        LSEBITax,
        LSBCTaxAmt,
        LMarAmt,
        LNStmAmt,
        LSGSTAmt,
        LTotBuyQty,
        LTotSellQty,
        GSTDChrs,
        LBrokAmount,
        LCGSTAmt,
        LTotBuyAmt,
        LTotSellAmt,
        LTotOpAmt,
        MSrvTax,
        LTranRate,
        LCBrokAmt,
        LcTranAmt,
        LSBCTax,
        LCGSTRate,
        LSGSTRate,
        LIGSTRate,
        LUTTRate,
        LClQty,
        LShareAmt,
        lOpQty,
        LIGSTAmt,
        LBalQty,
        LDiffAmt,
        LEQSTMRate,
        LEQSTTRate,
        LBrokType,
        LATranType,
        GSrTaxApp,
        LCttType,
        LRiskMType,
        LSEBIType,
        LCGSTType,
        LLotWise,
        LIGSTType,
        LUTTType,
        LPStmType,
        LMarType,
        LRiskMApp,
        MCLPattan,
        TRecVOU,
        LSGSTType,
        LCINNo,
        CountRec,
        LRptGrp,
        LPartyName,
        LContractAcc,
        LFAX,
        LPANNO,
        LPHoneo,
        LPhoneR,
        LPIN,
        LAExCode,
        LItemCode,
        LMobile,
        LInstType,
        LSaudaCode,
        LStmFlag,
        LOpFlag,
        Lmaturity,
        LSOpQty,
        LTillDate,
        LTFromDate,
        LSetNo,
        LBuyClRate,
        LSellClRate,
        LCloseAmt,
        LBillBy,
        LTotBQAmt,
        LTotSQAmt,
        LNetRecdPay,
        LNetMargin,
        LConType,
        LedgerFlag,
        AccStAdoREC,
        LOptCutBrok,
        LFutCutBrok,
        LNetPosition,
        LStampDutyDate,
        LNetOpenPosition,
        LOptType,
        LBrokName,
        Mledger1,
        Mledger2,
        Mledger3,
        LSaudaCode_muliple,
        loopdate,
        TempPart,
        mpartyop,
        MBROKTRADES,
        LPartyType,
        mpusdinr;

      const partiesRecordset = await adobeConnect(`
        SELECT DISTINCT A.ACCID,A.AC_CODE,A.NAME,A.OP_BAL,
        B.ADD1,B.CITY, B.PANNO, B.PARTYTYPE, B.PIN, B.PHONEO, B.PHONER, B.FAX, B.MOBILE, B.PIN, B.SRTAXAPP, B.CGST, B.SGST, B.IGST, B.UTT, B.CTTTYPE,B.RISKMTYPE,B.OPTCUTBROK,B.FUTCUTBROK,
        B.APPLYON,B.EMAIL,B.GSTIN,B.STATE,B.STATECODE,B.SEBITYPE,
        S.SAUDACODE,S.SAUDAID,S.INSTTYPE,I.ITEMCODE,S.STRIKEPRICE,S.OPTTYPE,S.MATURITY,S.TRADEABLELOT,S.BROKLOT ,S.REFLOT,
        I.ITEMID,I.LOT,I.RISKMAPP,I.SCGROUP,EX.EXCODE,EX.EXID,EX.LOTWISE,EX.CONTRACTACC
        FROM ACCOUNTM AS A, ACCOUNTD AS B, CTR_D AS C, ITEMMAST AS I, SAUDAMAST AS S ,EXMAST EX
        WHERE A.COMPCODE = ${req?.body?.companyCode} AND A.ACCID = B.ACCID
        AND S.ITEMID = I.ITEMID AND S.MATURITY >= '${
          req?.body?.fromDate
        }' AND C.CONDATE <= '${req?.body?.toDate}'
        AND A.ACCID = C.ACCID  AND C.SAUDAID = S.SAUDAID
        AND EX.EXID = S.EXID
        ${
          req?.body?.AllFmly === false &&
          req?.body?.LSFmlyIDs.length > 0 &&
          AllParties === true
            ? `AND A.ACCID IN (SELECT ACCID FROM ACCFMLYD WHERE FMLYID IN (${req?.body?.LSFmlyIDs}))`
            : ""
        }
        ${
          req?.body?.AllParties === false && req?.body?.LSParties.length > 0
            ? `AND A.AC_CODE IN (${req?.body?.LSParties})`
            : ""
        }
        ${
          req?.body?.AllExcodes === false && req?.body?.LSExCodes.length > 0
            ? `AND I.EXID IN (${req?.body?.LSExCodes})`
            : ""
        }
        ${
          req?.body?.AllSaudas === false
            ? `AND S.SAUDAID IN (${req?.body?.LSSaudas})`
            : ""
        }
        ${
          req?.body?.AllSaudas === true && req?.body?.AllInst === false
            ? `AND S.INSTTYPE IN (${req?.body?.LSInst})`
            : ""
        }
        ORDER BY A.NAME,EX.EXCODE,I.ITEMCODE,S.INSTTYPE,S.MATURITY
    `);

      // Party Loop Start //
      while (!partiesRecordset.EOF) {
        LPartyCode = partiesRecordset.Fields("AC_CODE").Value;
        LPartyName = partiesRecordset.Fields("NAME").Value;
        LACCID = partiesRecordset.Fields("ACCID").Value;
        LPANNO = (partiesRecordset.Fields("PANNO").Value || "").trim();
        LPHoneo = (partiesRecordset.Fields("PHONEO").Value || "").trim();
        LPhoneR = (partiesRecordset.Fields("PHONER").Value || "").trim();
        LFAX = (partiesRecordset.Fields("FAX").Value || "").trim();
        LMobile = (partiesRecordset.Fields("MOBILE").Value || "").trim();
        LPIN = (partiesRecordset.Fields("PIN").Value || "").trim();
        LCINNo = ""; // This seems to be always an empty string in your code
        LOptCutBrok = (
          partiesRecordset.Fields("OPTCUTBROK").Value || ""
        ).trim();
        LFutCutBrok = (
          partiesRecordset.Fields("FUTCUTBROK").Value || ""
        ).trim();
        LPartyType = partiesRecordset.Fields("PARTYTYPE").Value;

        if (req?.body?.GOnlyBrok === 0) {
          GSrTaxApp = partiesRecordset.Fields("SRTAXAPP").Value;
          LCttType =
            partiesRecordset.Fields("CTTTYPE").Value === null
              ? "N"
              : partiesRecordset.Fields("CTTTYPE").Value;
          LRiskMType =
            partiesRecordset.Fields("RISKMTYPE").Value === null
              ? "N"
              : partiesRecordset.Fields("RISKMTYPE").Value;
          LSEBIType =
            partiesRecordset.Fields("SEBITYPE").Value === null
              ? "N"
              : partiesRecordset.Fields("SEBITYPE").Value;
          LCGSTType =
            partiesRecordset.Fields("CGST").Value === null
              ? "N"
              : partiesRecordset.Fields("CGST").Value;
          LSGSTType =
            partiesRecordset.Fields("SGST").Value === null
              ? "N"
              : partiesRecordset.Fields("SGST").Value;
          LIGSTType =
            partiesRecordset.Fields("IGST").Value === null
              ? "N"
              : partiesRecordset.Fields("IGST").Value;
          LUTTType =
            partiesRecordset.Fields("UTT").Value === null
              ? "N"
              : partiesRecordset.Fields("UTT").Value;
          LPStmType =
            partiesRecordset.Fields("APPLYON").Value === null
              ? "N"
              : partiesRecordset.Fields("APPLYON").Value;
        }

        LTillDate = new Date(req?.body?.toDate);
        LTFromDate = new Date(req?.body?.fromDate);
        LMStrDate = "";

        LMStrDate = Get_MaxExCondate(
          partiesRecordset.Fields("EXID").Value,
          LTillDate
        );
        if (LMStrDate.length > 1) {
          LTillDate = new Date(LMStrDate);
        }
        LAExCode = partiesRecordset.Fields("EXCODE").Value;
        LItemCode = partiesRecordset.Fields("ITEMCODE").Value;
        LInstType = partiesRecordset.Fields("INSTTYPE").Value;
        LStrike = parseFloat(partiesRecordset.Fields("STRIKEPRICE").Value);
        LSaudaCode = partiesRecordset.Fields("SAUDACODE").Value;
        Lmaturity = partiesRecordset.Fields("MATURITY").Value;
        if (LSaudaCode_muliple === "") {
          LSaudaCode_muliple = LSaudaCode;
        } else {
          LSaudaCode_muliple = LSaudaCode_muliple + ";" + LSaudaCode;
        }

        const LRefLot =
          partiesRecordset.Fields("REFLOT").Value != null
            ? partiesRecordset.Fields("REFLOT").Value
            : 1;
        const LBrokLot =
          partiesRecordset.Fields("BROKLOT").Value != null
            ? partiesRecordset.Fields("BROKLOT").Value
            : 1;

        LContractAcc = partiesRecordset.Fields("CONTRACTACC").Value;
        LLotWise = partiesRecordset.Fields("LOTWISE").Value;
        LSaudaID = partiesRecordset.Fields("SAUDAID").Value;
        LExID = partiesRecordset.Fields("EXID").Value;
        LOptType = partiesRecordset.Fields("OPTTYPE").Value;
        LItemID = partiesRecordset.Fields("itemid").Value;
        LMinRate = 0;
        if (req?.body?.GMinBrokYN === "Y" && req?.body?.GFBroktype === "Y") {
          getminrate(LPartyCode, LItemID, LExID, LMinRate);
        }
        LBuyClRate = 1;
        LSellClRate = 1;

        if (req?.body?.LOpFlag === false) {
          LNetRecdPay = 0;

          // if (req?.body?.confirmed || req?.body?.all) {
          //   LprevBal = partiesRecordset.OP_BAL;
          //   if (req?.body?.req?.body?.Check10 === 0) {
          //     LBal = Net_DrCr(LPartyCode, req?.body?.fromDate);
          //     LprevBal += LBal;
          //   }
          //   if (req?.body?.confirmed) {
          //     if (req?.body?.showTradeConfirm) {
          //       LNetRecdPay = 0;
          //       LNetMargin = 0;
          //       TRec = null;
          //       let mysql = `SELECT SUM(CASE DR_CR  WHEN 'D' THEN AMOUNT * -1  WHEN 'C' THEN AMOUNT * 1 END) AS AMT
          //                    FROM VCHAMT
          //                    WHERE COMPCODE = ${body?.rec?.companyCode}  AND AC_CODE = '${partiesRecordset?.AC_CODE}' AND VOU_DT < '${body?.rec?.fromDate}' AND CHEQUE_NO NOT IN ('DMARGIN', 'RMARGIN');`;

          //       TRec = await adobeConnect(mysql);
          //       if (!TRec.EOF) LprevBal += IsNull(TRec.AMT, 0);

          //       mysql = `
          //     SELECT SUM(CASE DR_CR WHEN 'D' THEN AMOUNT * -1 WHEN 'C' THEN AMOUNT * 1 END) AS AMT
          //     FROM VCHAMT
          //     WHERE COMPCODE = ${req?.body?.companyCode}
          //     AND AC_CODE = '${LPartyCode}'
          //     AND VOU_DT >= '${Format(req?.body?.fromDate, "yyyy/MM/dd")}'
          //     AND VOU_DT <= '${Format(vcDTP2.Value, "yyyy/MM/dd")}'
          //     AND VOU_TYPE <> 'S'
          //     AND CHEQUE_NO NOT IN ('DMARGIN', 'RMARGIN')
          //              `;

          //       TRec = await adobeConnect(mysql);
          //       if (!TRec.EOF) LNetRecdPay = IsNull(TRec.AMT, 0);

          //       mysql = `
          //     SELECT IMARGIN
          //     FROM DLYCLMGN
          //     WHERE COMPCODE = ${req?.body?.companyCode}
          //       AND CLIENT = '${LPartyCode}'
          //       AND CONDATE = '${Format(vcDTP2.Value, "yyyy/MM/dd")}'
          //       AND EXCODE = 'CME'
          //   `;
          //       TRec = await adobeConnect(mysql);
          //       if (!TRec.EOF) LNetMargin = IsNull(TRec.IMARGIN, 0);

          //       TRec = null;
          //     } else {
          //       if (req?.body?.withSharing) {
          //         LprevBal += Get_Balance(LACCID);
          //       } else {
          //         TRec = null;
          //         let mysql = `EXEC ACSUMVCHS ${req?.body?.companyCode}, '${LPartyCode}', '${req?.body?.fromDate}', '${req?.body?.toDate}'`;
          //         TRec = await adobeConnect(mysql);
          //         if (!TRec.EOF) LprevBal += IsNull(TRec.AMT, 0);
          //         TRec = null;
          //       }
          //     }
          //   } else if (req?.body?.all) {
          //     LDebitAmt = 0;
          //     LCreditAmt = 0;
          //     let mysql = `
          //   SELECT DR_CR, SUM(AMOUNT) AS AMT
          //   FROM VCHAMT
          //   WHERE COMPCODE = ${req?.body?.companyCode}
          //     AND AC_CODE = '${LPartyCode}'
          //     AND VOU_DT >= '${Format(req?.body?.fromDate, "yyyy/MM/dd")}'
          //     AND VOU_TYPE <> 'S'
          //     AND VOU_DT <= '${Format(vcDTP2.Value, "yyyy/MM/dd")}'
          //   GROUP BY DR_CR
          // `;

          //     TRec = null;
          //     TRec = await adobeConnect(mysql);
          //     if (!TRec.EOF) {
          //       if (TRec.DR_CR === "C") LCreditAmt = IsNull(TRec.AMT, 0);
          //       if (TRec.DR_CR === "D") LDebitAmt = IsNull(TRec.AMT, 0);
          //       TRec.MoveNext();
          //       if (!TRec.EOF) {
          //         if (TRec.DR_CR === "C") LCreditAmt = IsNull(TRec.AMT, 0);
          //         if (TRec.DR_CR === "D") LDebitAmt = IsNull(TRec.AMT, 0);
          //       }
          //       LNetRecdPay = LCreditAmt - LDebitAmt;
          //     }
          //     TRec = null;
          //   }
          // }
          // LOpFlag = true;
        }

        while (
          LSaudaID === partiesRecordset.Fields("SAUDAID").Value &&
          LPartyCode === partiesRecordset.Fields("AC_CODE").Value
        ) {
          LTradeableLot = partiesRecordset.Fields("TRADEABLELOT").Value;

          if ((LAExCode = "NSE" && LLotWise == "Y")) {
            LCalval = partiesRecordset.Fields("TRADEABLELOT").Value;
          } else {
            if (LAExCode == "MCX" && LLotWise == "Y") {
              LCalval = partiesRecordset.Fields("TRADEABLELOT").Value;
            } else {
              LCalval = partiesRecordset.Fields("LOT").Value;
            }
          }
          mpusdinr = 1;
          console.log(
            `SELECT BILLBY,ACEXCODE FROM ACCT_EX WHERE COMPCODE = ${
              req?.body?.companyCode
            } AND AC_CODE = '${LPartyCode}' AND EXCODE = '${
              partiesRecordset.Fields("EXCODE").Value
            }'`
          );
          TRec2 = null;
          TRec2 = await adobeConnect(
            `SELECT BILLBY,ACEXCODE FROM ACCT_EX WHERE COMPCODE = ${
              req?.body?.companyCode
            } AND AC_CODE = '${LPartyCode}' AND EXCODE = '${
              partiesRecordset.Fields("EXCODE").Value
            }'`
          );
          if (!TRec2.EOF) {
            LBillBy = TRec2.Fields("BILLBY").Value;
          }
          TRec2 = null;
          console.log("billlll", LBillBy);
          if (LAExCode == "CMX") {
            TRec = null;
            TRec = await adobeConnect(
              `SELECT closerate,HGRATE,LOWRATE FROM CTR_R WHERE COMPCODE = ${
                req?.body?.companyCode
              } AND SAUDA = 'USDINR' AND CONDATE = '${Format(
                vcDTP2.Value,
                "YYYY/MM/DD"
              )}'`
            );
            if (!TRec.EOF) {
              LCurrFRate = TRec.CLOSERATE;
              LCurrCRate = TRec.LOWRATE;
              LCurrDRate = TRec.HGRATE;
            }
            TRec = null;
          } else {
            LCurrDRate = 1;
            LCurrCRate = 1;
          }

          if (LBillBy === "B") {
            let TRec = null;
            TRec = await adobeConnect(
              `SELECT CLOSERATE, HGRATE, LOWRATE FROM CTR_R WHERE COMPCODE = ${
                req?.body?.companyCode
              } AND SAUDAID = '${LSaudaID}' AND CONDATE = '${Format(
                vcDTP2.Value,
                "yyyy/mm/dd"
              )}'`
            );
            if (!TRec.EOF) {
              LBuyClRate = TRec.LOWRATE === 0 ? TRec.CLOSERATE : TRec.LOWRATE;
              LSellClRate = TRec.HGRATE === 0 ? TRec.CLOSERATE : TRec.HGRATE;
            }
          } else if (LBillBy === "S") {
            LCurrCRate = LCurrFRate;
            LCurrDRate = LCurrFRate;
          } else if (LBillBy === "P") {
            LPOPRATE = 0;
            let LPCLRATE = 0;
            TRec = null;
            TRec = await adobeConnect(
              `SELECT TOP 1 SETTLERATE FROM CTR_RP WHERE party = '${LPartyCode}' AND SAUDAID = '${LSaudaID}' AND CONDATE < '${Format(
                req?.body?.fromDate,
                "yyyy/mm/dd"
              )}' ORDER BY CONDATE DESC`
            );
            if (!TRec.EOF) {
              if (TRec.SETTLERATE !== 0) LPOPRATE = TRec.SETTLERATE;
            }

            TRec = await adobeConnect(
              `SELECT SETTLERATE FROM CTR_RP WHERE PARTY = '${LPartyCode}' AND SAUDAID = '${LSaudaID}' AND CONDATE = '${Format(
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
              } AND PARTY = '${LPartyCode}' AND SAUDA = 'USDINR' AND CONDATE = '${Format(
                vcDTP2.Value,
                "YYYY/MM/DD"
              )}'`
            );
            if (!TRec.EOF) {
              LCurrFRate = TRec.SETTLERATE;
              LCurrCRate = TRec.LOW;
              LCurrDRate = TRec.HIGH;
            }
          }

          // chek check 13 //
          if (req?.body?.GUniqClientId === "BRO2-CHE") {
            if (
              mpusdinr === 1 &&
              LAExCode === "CMX" &&
              Check5.Value === 0 &&
              Check13.Value === 1
            ) {
              let TRec = null;
              TRec = await adobeConnect(
                `SELECT CURRRATE, USDINR FROM INV_D WHERE COMPCODE = ${
                  req?.body?.companyCode
                } AND PARTY = '${LPartyCode}' AND EXCODE = 'CMX' AND STDATE >= '${Format(
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
              LAExCode === "CMX" &&
              Check5.Value === 0 &&
              Check13.Value === 1
            ) {
              let TRec = null;
              TRec = await adobeConnect(
                `SELECT CURRRATE, USDINR FROM INV_D WHERE COMPCODE = ${
                  req?.body?.companyCode
                } AND PARTY = '${LPartyCode}' AND EXCODE = 'CMX' AND STDATE >= '${Format(
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
          if (LAExCode === "CMX") {
            // LCurrCRate = 1;
            // LCurrDRate = 1;
          }
          LRiskMApp = partiesRecordset.RISKMAPP;
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
            )}','${formatDateYYYYMMDD(
              LTillDate
            )}',${LSaudaID},'${LPartyCode}' ,${req?.body?.GOnlyBrok}`
          );

          if (req?.body?.GOnlyBrok == 0) {
            if (LPStmType == "R" || LPStmType == "P") {
              if (LTillDate >= GSTMDate && LTillDate < LStampDutyDate) {
                LNStmAmt = Get_StampDutyAmt(
                  LPartyCode,
                  LTFromDate,
                  LTillDate,
                  LSExCodes
                );
              }
            }
          }

          if (req?.body?.Check10 == 1) {
            TRec = null;
            let TRec = await adobeConnect(
              `SELECT A.ROWNO1, A.CONNO, A.QTY, A.RATE, A.CONTYPE, A.CONDATE, A.pattan 
               FROM CTR_D AS A 
               WHERE A.COMPCODE = ${req?.body?.companyCode} AND A.AC_CODE = ${LPartyCode} AND A.CONDATE BETWEEN ${req?.body?.fromDate} AND '${req?.body?.toDate}' AND A.SAUDAID = ${LSaudaID}
               ORDER BY A.CONDATE, A.CONNO`
            );
            while (!TRec.EOF) {
              CountRec++;
              AccRecSet.AddNew();

              AccRecSet.Fields("RPTGRP").Value = 1;
              AccRecSet.Fields("SRNO").Value = CountRec;
              AccRecSet.Fields("CALVAL").Value = LCalval;
              AccRecSet.Fields("CONNO").Value = TRec.Fields("CONNO").Value;

              if (TRec.Fields("CONTYPE").Value > "B") {
                AccRecSet.Fields("CTYPE").Value = "B";
                AccRecSet.Fields("CONTDATE").Value = TRec.Fields("Condate")
                  .Value.toISOString()
                  .slice(0, 10);
                AccRecSet.Fields("BQNTY").Value = TRec.Fields("QTY").Value;
                AccRecSet.Fields("BRATE").Value = TRec.Fields("Rate").Value;
                AccRecSet.Fields("DRAMOUNT").Value = 0;
              } else {
                AccRecSet.Fields("CTYPE").Value = "S";
                AccRecSet.Fields("CONTDATE").Value = TRec.Fields("Condate")
                  .Value.toISOString()
                  .slice(0, 10);
                AccRecSet.Fields("SQNTY").Value = TRec.Fields("QTY").Value;
                AccRecSet.Fields("SRATE").Value = TRec.Fields("Rate").Value;
                AccRecSet.Fields("DRAMOUNT").Value = 0;
              }

              Add_PartyDetails();
              Add_OtherDetails();
              AccRecSet.Update();
              TRec.MoveNext();
            }

            TRec.Close();
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

          lOpQty = SDOpQty(LPartyCode, LSaudaID, LTFromDate);
          lOpQty = Math.round(lOpQty * 100) / 100; // Round to 2 decimal places
          LNetOpenPosition = lOpQty;
          LNetPosition = lOpQty;
          LBalQty = lOpQty;

          if (lOpQty !== 0) {
            LOpenRate = SDCLRATE(LSaudaID, LTFromDate - 1, "O"); // closing rate for opening
            if (LPOPRATE !== 0) {
              LOpenRate = LPOPRATE;
            }
            LTotOpAmt = lOpQty * LOpenRate * LCalval;
          }

          if (
            (LInstType === "OPT" && OptMTMChk.Value === 0) ||
            (LInstType === "CSH" && CshMTMchk.Value === 0)
          ) {
            lOpQty = 0;
            LTotOpAmt = 0;
          }

          if (lOpQty !== 0) {
            if (req?.body?.MFormat === "Account Statement") {
              CountRec++;
              let LConType;
              let LGBrokName;
              let LGBrokRate;
              let LRptGrp;

              if (lOpQty > 0) {
                LConType = "B";
                LDiffAmt = Math.abs(lOpQty) * LOpenRate * LCalval * -1;
              } else {
                LConType = "S";
                LDiffAmt = Math.abs(lOpQty) * LOpenRate * LCalval;
              }

              if (req?.body?.Check10 === 1) {
                LRptGrp = 2;
              } else {
                LRptGrp = 1;
              }

              switch (LBrokType) {
                case "P":
                  LGBrokName = "Percentage Wise";
                  break;
                case "T":
                  LGBrokName = "Transaction Wise";
                  break;
                case "O":
                  LGBrokName = "Opening Sauda";
                  break;
              }
              LGBrokRate = String(LBrokRate.toFixed(4));

              if (
                !(req?.body?.Check12 === 1 && MFormat === "Account Statement")
              ) {
                Add_To_AccRecSet(
                  LRptGrp,
                  CountRec,
                  LSetNo,
                  lOpQty,
                  LOpenRate,
                  "",
                  LTFromDate - 1,
                  LNetRecdPay,
                  LNetMargin,
                  "",
                  0,
                  0,
                  LprevBal,
                  LConType,
                  LRefLot,
                  LCalval,
                  "O",
                  "*",
                  LInstType
                );
              }
            }
          }

          if (
            req?.body?.Check12 === 1 &&
            !mpartyop &&
            MFormat === "Account Statement"
          ) {
            if (
              req?.body?.Check10 === 1 &&
              Option5.Value &&
              req?.body?.Check12 === 1 &&
              MFormat === "Account Statement"
            ) {
              let LBal = Get_ClosingBal(
                partiesRecordset.ACCID,
                req?.body?.fromDate - 1
              );
              mpartyop = true;
              if (LBal !== 0) {
                CountRec++;
                Add_VOUCHER(
                  "2",
                  CountRec,
                  "OPENING",
                  req?.body?.fromDate,
                  "",
                  "",
                  "",
                  0,
                  LBal
                );
              }
            }
          }

          // Bought incertion //
          TRec = null;
          TRec = new ActiveXObject("ADODB.Recordset");
          TRec = AccStAdoREC;

          if (req?.body.GCINNo === "2000") {
            TRec.Filter =
              "CONDATE >= '" + LTFromDate.toISOString().slice(0, 10) + "'";
          } else {
            if (req?.body?.Check12 === 0) {
              TRec.Filter = "B'";
            }
          }

          if (!TRec.EOF && loopdate === "01/01/1900") {
            loopdate = TRec.Fields("Condate").Value;
          }

          while (!TRec.EOF) {
            LCBrokAmt = 0;
            LcTranAmt = 0;

            LBrokType = TRec.Fields("BROKTYPE").Value
              ? TRec.Fields("BROKTYPE").Value
              : "T";
            LBrokRate2 = parseFloat(TRec.Fields("BROKRATE2").Value) || 0;
            LBrokQty =
              TRec.Fields("BROKQTY").Value != null
                ? TRec.Fields("BROKQTY").Value
                : 0;
            LBrokRate = parseFloat(TRec.Fields("BROKRATE").Value) || 0;
            LBrokQty2 =
              TRec.Fields("BROKQTY2").Value != null
                ? TRec.Fields("BROKQTY2").Value
                : 0;

            if (TRec.Fields("ORDNO").Value != null) {
              if (TRec.Fields("ORDNO").Value !== "Carry") {
                LBrokRate = LBrokRate2;
              }
            }

            if (req?.body?.GOnlyBrok === 0) {
              LTranRate = parseFloat(TRec.Fields("TRANRATE").Value) || 0;
              LATranType =
                TRec.Fields("TRANTYPE").Value != null
                  ? TRec.Fields("TRANTYPE").Value
                  : "T";
              MSrvTax = parseFloat(TRec.Fields("SRVTAX").Value) || 0;
              LSBCTax = parseFloat(TRec.Fields("SBC_TAX").Value) || 0;
              LCGSTRate = parseFloat(TRec.Fields("CGSTRATE").Value) || 0;
              LSGSTRate = parseFloat(TRec.Fields("SGSTRATE").Value) || 0;
              LIGSTRate = parseFloat(TRec.Fields("IGSTRATE").Value) || 0;
              LUTTRate = parseFloat(TRec.Fields("UTTRATE").Value) || 0;
              LSEBITax = parseFloat(TRec.Fields("SEBITAX").Value) || 0;
              LSTTRate = parseFloat(TRec.Fields("STTRATE").Value) || 0;
              LEQSTTRate = parseFloat(TRec.Fields("EQ_STT").Value) || 0;
              LEQSTMRate = parseFloat(TRec.Fields("EQ_STAMP").Value) || 0;

              if (
                TRec.Fields("CONDATE").Value >= GSTMDate &&
                TRec.Fields("CONDATE").Value < LStampDutyDate &&
                (LPStmType === "R" || LPStmType === "P")
              ) {
                LStmRate = 0;
              } else {
                LStmRate = parseFloat(TRec.Fields("STMRATE").Value) || 0;
                LEQSTMRate = parseFloat(TRec.Fields("EQ_STAMP").Value) || 0;
              }
            }

            if (TRec.Fields("PATTAN").Value !== "C") {
              if (req?.body?.GCINNo === "2000") {
                if (TRec.Fields("CONTYPE").Value === "B") {
                  lOpQty += TRec.Fields("QTY").Value;
                  LTotOpAmt +=
                    TRec.Fields("QTY").Value *
                    TRec.Fields("RATE").Value *
                    LCalval;
                  LBalQty = lOpQty;
                } else {
                  lOpQty -= TRec.Fields("QTY").Value;
                  LTotOpAmt -=
                    TRec.Fields("QTY").Value *
                    TRec.Fields("RATE").Value *
                    LCalval;
                  LBalQty = lOpQty;
                }
              } else {
                if (!CONTYPE === "B") {
                  lOpQty += TRec.Fields("QTY").Value;
                  LTotOpAmt +=
                    TRec.Fields("QTY").Value *
                    TRec.Fields("RATE").Value *
                    LCalval;
                  LBalQty = lOpQty;
                } else {
                  lOpQty -= TRec.Fields("QTY").Value;
                  LTotOpAmt -=
                    TRec.Fields("QTY").Value *
                    TRec.Fields("RATE").Value *
                    LCalval;
                  LBalQty = lOpQty;
                }
              }
              LNetPosition = lOpQty;
            } else {
              if (req?.body?.GCINNo === "2000") {
                if (TRec.Fields("CONTYPE").Value === "B") {
                  LBuyAmt =
                    TRec.Fields("QTY").Value *
                    TRec.Fields("RATE").Value *
                    LCalval;
                  LTotBuyAmt += LBuyAmt;
                } else {
                  LSellAmt =
                    TRec.Fields("QTY").Value *
                    TRec.Fields("RATE").Value *
                    LCalval;
                  LTotSellAmt += LSellAmt;
                }
              } else {
                LBuyAmt =
                  TRec.Fields("QTY").Value *
                  TRec.Fields("RATE").Value *
                  LCalval;
                LTotBuyAmt += LBuyAmt;
                LTotBQAmt += LBuyAmt;
              }

              if (LBrokType === "Z") LBrokQty = LBrokLot;
              if (LBrokType === "A") LBrokQty2 = LBrokLot;
              if (LBrokType === "R") LBrokQty2 = LBrokLot;
              if (LBrokType === "5") LBrokQty2 = LBrokLot;

              if (req?.body?.GCINNo === "2000" || req?.body?.Check12 === 1) {
                if (!CONTYPE === "B") {
                  if (LBrokType === "O" || LBrokType === "C")
                    LBrokQty = Get_BrokQty(
                      LBrokType,
                      LBalQty,
                      "B",
                      TRec.Fields("QTY").Value
                    );
                  LBalQty += TRec.Fields("QTY").Value;
                  LTotBuyQty += TRec.Fields("QTY").Value;
                } else {
                  if (LBrokType === "O" || LBrokType === "C")
                    LBrokQty = Get_BrokQty(
                      LBrokType,
                      LBalQty,
                      "S",
                      TRec.Fields("QTY").Value
                    );
                  LBalQty += TRec.Fields("QTY").Value;
                  LTotSellQty += TRec.Fields("QTY").Value;
                }
              } else {
                if (
                  LBrokType === "O" ||
                  LBrokType === "C" ||
                  LBrokType === "A" ||
                  LBrokType === "5"
                )
                  LBrokQty = Get_BrokQty(
                    LBrokType,
                    LBalQty,
                    "B",
                    TRec.Fields("QTY").Value
                  );
                LBalQty += TRec.Fields("QTY").Value;
                LTotBuyQty += TRec.Fields("QTY").Value;
              }
              LCBrokAmt = calcBrokerage(
                LBrokType,
                LBrokRate,
                LBrokRate2,
                LBrokQty,
                TRec.Fields("QTY").Value,
                TRec.Fields("RATE").Value,
                LCalval,
                LInstType,
                LStrike,
                LBrokQty2,
                LTradeableLot,
                "B",
                "2",
                LItemID
              );

              if (LMinRate !== 0 && LMinRate * (!QTY / LRefLot) > LCBrokAmt) {
                LCBrokAmt = LMinRate * (!QTY / LRefLot);
              }

              LBrokAmount += LCBrokAmt;

              if (req?.body?.GOnlyBrok === 0) {
                if (LInstType === "CSH") {
                  LStmAmt =
                    (LStmRate / 100) * ((!QTY - LBrokQty2) * !Rate * LCalval);
                  LTotStmAmt += LStmAmt;
                  LTotStmAmt +=
                    (LEQSTMRate / 100) * (LBrokQty2 * !Rate * LCalval);
                  LSTTAmt =
                    (LSTTRate / 100) * ((!QTY - LBrokQty2) * !Rate * LCalval);
                  LTotSTTAmt += LSTTAmt;
                  LTotSTTAmt +=
                    (LEQSTTRate / 100) * (LBrokQty2 * !Rate * LCalval);
                } else {
                  LStmAmt = (LStmRate / 100) * LBuyAmt;
                  LSTTAmt = (LSTTRate / 100) * LBuyAmt;
                  LTotStmAmt += LStmAmt;
                  LTotSTTAmt += LSTTAmt;
                }

                LSEBITaxAmt = (LSEBITax / 100) * LBuyAmt;
                LTotSEBITaxAmt += LSEBITaxAmt;

                if (LATranType === "T") {
                  LcTranAmt = LTranRate * !QTY;
                  LTranAmount += LcTranAmt;
                } else if (LATranType === "P") {
                  LcTranAmt = (LTranRate / 100) * LBuyAmt;
                  LTranAmount += LcTranAmt;
                } else if (LATranType === "I") {
                  LcTranAmt = (LTranRate / 100) * LBrokQty * !Rate * LCalval;
                  LTranAmount += LcTranAmt;
                }

                GSrvTax += (LCBrokAmt * MSrvTax) / 100;
                LSBCTaxAmt += (LCBrokAmt * LSBCTax) / 100;
                LCGSTAmt += (LCBrokAmt * LCGSTRate) / 100;
                LSGSTAmt += (LCBrokAmt * LSGSTRate) / 100;
                LIGSTAmt += (LCBrokAmt * LIGSTRate) / 100;
                LUTTAmt += (LCBrokAmt * LUTTRate) / 100;

                if (GSrStdYN === "Y") {
                  if (!CONDATE >= DateValue(GSrvDate)) {
                    GSrvTax += (LcTranAmt * MSrvTax) / 100;
                    LSBCTaxAmt += (LcTranAmt * LSBCTax) / 100;
                    LCGSTAmt += (LcTranAmt * LCGSTRate) / 100;
                    LSGSTAmt += (LcTranAmt * LSGSTRate) / 100;
                    LIGSTAmt += (LcTranAmt * LIGSTRate) / 100;
                    LUTTAmt += (LcTranAmt * LUTTRate) / 100;
                  }
                  if (TRec.Fields("Condate").Value >= DateValue(GSrvDate2)) {
                    LCGSTAmt += (LSEBITaxAmt * LCGSTRate) / 100;
                    LSGSTAmt += (LSEBITaxAmt * LSGSTRate) / 100;
                    LIGSTAmt += (LSEBITaxAmt * LIGSTRate) / 100;
                  }
                }
              }

              if (req?.body?.MFormat === "Account Statement") {
                CountRec++;
                if (
                  LBrokType === "P" ||
                  LBrokType === "T" ||
                  LBrokType === "O"
                ) {
                  LGBrokName =
                    LBrokType === "P"
                      ? "Percentage Wise"
                      : LBrokType === "T"
                      ? "Transaction Wise"
                      : "Opening Sauda";
                  LGBrokRate = LBrokRate.toFixed(4);
                  if (!(PATTAN !== "C" && req?.body?.Check12 === 1)) {
                    Add_To_AccRecSet(
                      LRptGrp,
                      CountRec,
                      LSetNo,
                      !QTY,
                      !Rate,
                      !ROWNO1,
                      !Condate,
                      LNetRecdPay,
                      LNetMargin,
                      !contime,
                      LBrokRate,
                      LCBrokAmt,
                      LprevBal,
                      !CONTYPE,
                      LRefLot,
                      LCalval,
                      !PATTAN,
                      LBrokType,
                      LInstType
                    );
                  }
                }
              }
            }
            TRec.MoveNext();
          }

          CFlag = false;

          // End Bought incertion //

          //Add converted code hear

          if (req?.body?.Check12 === 0) {
            if (req?.body?.GCINNo !== "2000") {
              let TRec = await AccStAdoREC.clone();
              TRec.filter = "'S'";
              await TRec.moveFirst();

              while (!TRec.EOF) {
                console.log(
                  "LBrokQty2====================================>",
                  LBrokQty2
                );
                AccRecSet.filter = adFilterNone;
                AccRecSet.sort = "SRNO ASC";
                AccRecSet.filter =
                  "'" +
                  LSaudaCode +
                  "' AND PARTYCODE='" +
                  LPartyCode +
                  "' And PATTAN ='C'";
                await AccRecSet.moveFirst();

                while (!AccRecSet.EOF) {
                  let LCBrokAmt = 0,
                    LcTranAmt = 0;
                  let LBrokType = TRec.broktype === null ? "T" : TRec.broktype;
                  let LBrokRate = parseFloat(TRec.brokrate);
                  let LBrokRate2 = parseFloat(TRec.BROKRATE2);
                  let LBrokQty = TRec.BROKQTY === null ? 0 : TRec.BROKQTY;
                  let LBrokQty2 = TRec.BROKQTY2 === null ? 0 : TRec.BROKQTY2;

                  if (TRec.ORDNO !== null && TRec.ORDNO === "Carry") {
                    LBrokRate = LBrokRate2;
                  }

                  if (GOnlyBrok === 0) {
                    let LATranType =
                      TRec.TRANTYPE === null ? "T" : TRec.TRANTYPE;
                    let LTranRate = parseFloat(TRec.TRANRATE);

                    if (
                      TRec.Condate >= GSTMDate &&
                      TRec.Condate < LStampDutyDate &&
                      (LPStmType === "R" || LPStmType === "P")
                    ) {
                      LStmRate = 0;
                    } else {
                      LStmRate = parseFloat(TRec.STMRATE);
                    }

                    LStmRate = parseFloat(TRec.STMRATE);
                    LEQSTTRate = parseFloat(TRec.EQ_STT);
                    LEQSTMRate = parseFloat(TRec.EQ_STAMP);
                    LSTTRate = parseFloat(TRec.STTRATE);
                    LSEBITax = parseFloat(TRec.SEBITAX);
                    MSrvTax = parseFloat(TRec.SRVTAX);
                    LSBCTax = parseFloat(TRec.SBC_TAX);
                    LCGSTRate = parseFloat(TRec.CGSTRATE);
                    LSGSTRate = parseFloat(TRec.SGSTRATE);
                    LIGSTRate = parseFloat(TRec.IGSTRATE);
                    LUTTRate = parseFloat(TRec.UTTRATE);
                    LSEBITax = parseFloat(TRec.SEBITAX);
                    LSTTRate = parseFloat(TRec.STTRATE);
                  }

                  if (TRec.PATTAN !== "C") {
                    lOpQty -= TRec.QTY;
                    LTotOpAmt -= TRec.QTY * TRec.Rate * LCalval;
                    LCBrokAmt = 0;
                    LNetPosition -= Math.abs(TRec.QTY);
                  } else {
                    LSellAmt = TRec.QTY * TRec.Rate * LCalval;
                    LTotSellAmt += LSellAmt;
                    LTotSellQty += TRec.QTY;
                    LTotSQAmt += LSellAmt;
                    LBrokAmount += LCBrokAmt;
                    if (GOnlyBrok === 0) {
                      if (LInstType === "CSH") {
                        LStmAmt =
                          (LStmRate / 100) *
                          ((TRec.QTY - LBrokQty2) * TRec.Rate * LCalval);
                        LTotStmAmt += LStmAmt;
                        LTotStmAmt +=
                          (LEQSTMRate / 100) *
                          (LBrokQty2 * TRec.Rate * LCalval);
                        LSTTAmt =
                          (LSTTRate / 100) *
                          ((TRec.QTY - LBrokQty2) * TRec.Rate * LCalval);
                        LTotSTTAmt += LSTTAmt;
                        LTotSTTAmt +=
                          (LEQSTTRate / 100) *
                          (LBrokQty2 * TRec.Rate * LCalval);
                      } else {
                        LStmAmt = (LStmRate / 100) * LSellAmt;
                        LSTTAmt = (LSTTRate / 100) * LSellAmt;
                        LTotStmAmt += LStmAmt;
                        LTotSTTAmt += LSTTAmt;
                      }
                      LSEBITaxAmt = (LSEBITax / 100) * LSellAmt;
                      LTotSEBITaxAmt += LSEBITaxAmt;
                      GSrvTax += (LCBrokAmt * MSrvTax) / 100;
                      LSBCTaxAmt += (LCBrokAmt * LSBCTax) / 100;
                      LcTranAmt = 0;
                      if (LATranType === "P") {
                        LcTranAmt =
                          (LTranRate / 100) * TRec.QTY * TRec.Rate * LCalval;
                        LTranAmount += LcTranAmt;
                      }
                      LCGSTRate = parseFloat(TRec.CGSTRATE);
                      LSGSTRate = parseFloat(TRec.SGSTRATE);
                      LIGSTRate = parseFloat(TRec.IGSTRATE);
                      LUTTRate = parseFloat(TRec.UTTRATE);
                      LCGSTAmt += (LCBrokAmt * LCGSTRate) / 100;
                      LSGSTAmt += (LCBrokAmt * LSGSTRate) / 100;
                      LIGSTAmt += (LCBrokAmt * LIGSTRate) / 100;
                      LUTTAmt += (LCBrokAmt * LUTTRate) / 100;

                      if (GSrStdYN === "Y") {
                        if (TRec.Condate >= DateValue(GSrvDate)) {
                          GSrvTax += (LcTranAmt * MSrvTax) / 100;
                          LSBCTaxAmt += (LcTranAmt * LSBCTax) / 100;
                          LCGSTAmt += (LcTranAmt * LCGSTRate) / 100;
                          LSGSTAmt += (LcTranAmt * LSGSTRate) / 100;
                          LIGSTAmt += (LcTranAmt * LIGSTRate) / 100;
                          LUTTAmt += (LcTranAmt * LUTTRate) / 100;
                        }
                        if (TRec.Condate >= DateValue(GSrvDate2)) {
                          LCGSTAmt += (LSEBITaxAmt * LCGSTRate) / 100;
                          LSGSTAmt += (LSEBITaxAmt * LSGSTRate) / 100;
                          LIGSTAmt += (LSEBITaxAmt * LIGSTRate) / 100;
                        }
                      }
                    }
                  }

                  AccRecSet.moveNext();
                }
                TRec.moveNext();
              }
            }
          }
          LCBrokAmt = 0;

          // CLOSING CALC
          if (
            req?.body?.MFormat === "Bill Summary" ||
            req?.body?.MFormat === "Bill Summary With Sharing" ||
            req?.body?.MFormat === "Account Statement Summary"
          ) {
            if (LTotOpAmt > 0) {
              LTotBuyAmt += LTotOpAmt;
            } else {
              LTotSellAmt += Math.abs(LTotOpAmt);
            }
          }

          // LClQty calculation
          LClQty =
            LNetPosition + parseFloat((LTotBuyQty - LTotSellQty).toFixed(2));
          console.log("LClQty=======================>", LClQty);
          if (
            LInstType === "OPT" &&
            OptMTMChk.Value === 0 &&
            Lmaturity <= LTillDate
          ) {
            LClQty = parseFloat(
              LTotBuyQty - LTotSellQty + -1 * parseFloat(LNetOpenPosition)
            );
          }

          if (
            (LInstType === "OPT" && OptMTMChk.Value === 0) ||
            (LInstType === "CSH" && CshMTMchk.Value === 0)
          ) {
            LClQty = 0;
          }

          // Update LNetPosition
          LNetPosition += parseFloat((LTotBuyQty - LTotSellQty).toFixed(2));

          // Calculate LCloseRate
          if (LBillBy === "P") {
            if (LPCLRATE !== 0) {
              LCloseRate = LPCLRATE;
            } else {
              if (LClQty !== 0) {
                LCloseRate = SDCLRATE(
                  LSaudaID,
                  Lmaturity <= LTillDate ? Lmaturity : LTillDate,
                  "C"
                ); // CLOSING RATE
              }
            }
          } else {
            if (LClQty !== 0) {
              LCloseRate = SDCLRATE(
                LSaudaID,
                Lmaturity <= LTillDate ? Lmaturity : LTillDate,
                "C"
              ); // CLOSING RATE
            }
          }
          LMarAmt = 0;

          if (
            (LInstType === "OPT" && OptMTMChk.Value === 0) ||
            (LInstType === "CSH" && CshMTMchk.Value === 0)
          ) {
            LClQty = 0;
          }

          if (
            LInstType === "OPT" &&
            OptMTMChk.Value === 0 &&
            Lmaturity <= LTillDate &&
            LCloseRate !== 0
          ) {
            LClQty = parseFloat(
              LTotBuyQty - LTotSellQty + parseFloat(LNetOpenPosition)
            );

            if (LOptType === "CE" && LCloseRate >= LStrike) {
              LCloseRate = parseFloat((LCloseRate - LStrike).toFixed(2));
            } else if (LOptType === "CE" && LCloseRate < LStrike) {
              LClQty = 0;
              LCloseRate = 0;
            }

            if (LOptType === "PE" && LCloseRate > LStrike) {
              LCloseRate = 0;
              LClQty = 0;
            }
          }
          if (req?.body?.Option1 === true) {
            if (LTillDate < Lmaturity) {
              if (LClQty !== 0) {
                LMarRate = 0;
                LMarType = "V";
                TRec = null;
                TRec = new ADODB.Recordset(); // BROKERAGE CALCULATION
                mysql =
                  "SELECT TOP 1 MARTYPE, MARRATE FROM PITBROK WHERE COMPCODE = " +
                  req?.body?.companyCode +
                  " And UPTOSTDT >= '" +
                  Format(LTillDate, "yyyy/MM/dd") +
                  "' AND AC_CODE='" +
                  LPartyCode +
                  "' AND ITEMID=" +
                  LItemID +
                  " AND INSTTYPE ='" +
                  LInstType +
                  "' ORDER BY UPTOSTDT ";
                TRec.Open(mysql, Cnn, adOpenForwardOnly, adLockReadOnly);
                if (TRec.EOF) {
                  TRec = null;
                  TRec = new ADODB.Recordset(); // BROKERAGE CALCULATION
                  mysql =
                    "SELECT TOP 1 MARTYPE, MARRATE FROM PEXBROK WHERE COMPCODE = " +
                    req?.body?.companyCode +
                    " And UPTOSTDT >= '" +
                    Format(LTillDate, "yyyy/MM/dd") +
                    "' AND AC_CODE='" +
                    LPartyCode +
                    "' AND EXID =" +
                    LExID +
                    " AND INSTTYPE ='" +
                    LInstType +
                    "' ORDER BY UPTOSTDT ";
                  TRec.Open(mysql, Cnn, adOpenForwardOnly, adLockReadOnly);
                  if (TRec.EOF) {
                    LMarRate = 0;
                    LMarType = "";
                    LMarRate = 0;
                  } else {
                    LMarType = IsNull(TRec.MARTYPE) ? "T" : TRec.MARTYPE;
                    LMarRate = IsNull(TRec.MARRATE) ? 0 : TRec.MARRATE;
                    LMarRate = LMarRate;
                  }
                } else {
                  LMarType = IsNull(TRec.MARTYPE) ? "T" : TRec.MARTYPE;
                  LMarRate = IsNull(TRec.MARRATE) ? 0 : TRec.MARRATE;
                }
                TRec = null;
                LMarAmt = Calc_Margin(
                  LMarType,
                  LMarRate,
                  LSaudaCode,
                  LTillDate,
                  Math.abs(LClQty),
                  LCloseRate,
                  LPartyCode,
                  LCalval,
                  LAExCode
                );
              } // LClQty
            }
          }
          if (LPartyCode === LContractAcc) LMarAmt *= -1;

          if (LClQty !== 0 && LCloseRate === 0)
            if (
              MsgBox(
                "Closing Rate missing For Sauda " +
                  LSaudaCode +
                  " of Date " +
                  CStr(LTillDate),
                vbOKCancel
              ) === vbCancel
            )
              return;

          MCLPattan = "B";

          if (LTillDate >= Lmaturity && LClQty !== 0 && LBrokType !== "3") {
            LBrokRate = 0;
            LBrokType = "P";
            LATranType = "P";
            LTranRate = 0;
            LBrokRate2 = 0;

            if (LInstType === "OPT" && LOptCutBrok === "0") {
              LBrokRate = 0;
              LBrokType = "T";
              LBrokRate2 = 0;
              LATranType = "P";
              LTranRate = 0;
            } else if (LInstType === "FUT" && LFutCutBrok === "0") {
              LBrokRate = 0;
              LBrokType = "T";
              LBrokRate2 = 0;
              LATranType = "P";
              LTranRate = 0;
            } else {
              TRec = null;
              TRec = new ADODB.Recordset(); // BROKERAGE CALCULATION
              mysql = `" EXEC GET_BROKRATE " + LACCID + "," + LExID + "," + LItemID + '" & Format(Lmaturity, "YYYY/MM/DD") & "','" & LInstType & "'";`;
              TRec.Open(mysql, Cnn, adOpenForwardOnly, adLockReadOnly);

              if (!TRec.EOF) {
                if (TRec.MinRate > LCloseRate) {
                  LBrokRate = TRec.MBROKRATE;
                  LBrokType = TRec.MBROKTYPE;
                  LBrokRate2 = TRec.MBROKRATE2;
                  LATranType = TRec.TRANTYPE;
                  LTranRate = TRec.TRANRATE;
                } else {
                  LBrokRate = TRec.brokrate;
                  LBrokType = TRec.broktype;
                  LBrokRate2 = TRec.BROKRATE2;
                  LATranType = TRec.TRANTYPE;
                  LTranRate = TRec.TRANRATE;
                }
              }
              TRec = null;
            }

            MSrvTax = 0;
            if (
              GSrTaxApp === "1" ||
              LCGSTType === "1" ||
              LIGSTType === "1" ||
              LIGSTType === "1"
            ) {
              TRec = null;
              TRec = new ADODB.Recordset();
              mysql =
                "SELECT SERVICETAX, SBC_TAX, STCTAX, CGSTRATE, SGSTRATE, IGSTRATE, UTTRATE FROM EXTAX WHERE COMPCODE =" +
                req?.body?.companyCode +
                " AND EXCHANGECODE = '" +
                LAExCode +
                "' AND FROMDT <= '" +
                Format(Lmaturity, "YYYY/MM/DD") +
                "' AND TODT >= '" +
                Format(Lmaturity, "YYYY/MM/DD") +
                "'";
              TRec.Open(mysql, Cnn, adOpenForwardOnly, adLockReadOnly);

              if (!TRec.EOF) {
                MSrvTax = Val(TRec.servicetax);
                LSBCTax = Val(TRec.SBC_TAX);
                LSEBITax = Val(TRec.STCTAX);
                LCGSTRate = Val(TRec.CGSTRATE);
                LSGSTRate = Val(TRec.SGSTRATE);
                LIGSTRate = Val(TRec.IGSTRATE);
                LUTTRate = Val(TRec.UTTRATE);
              } else {
                MSrvTax = 0;
                LSBCTax = 0;
                LCGSTRate = 0;
                LSBCTax = 0;
                LSGSTRate = 0;
                LIGSTRate = 0;
                LUTTRate = 0;
                LSEBITax = 0;
              }
              TRec = null;
              TRec = new ADODB.Recordset();
              mysql =
                "SELECT SEBITAX FROM ITEM_TAX WHERE COMPCODE =" +
                req?.body?.companyCode +
                " AND ITEMID =" +
                LItemID +
                " AND '" +
                Format(Lmaturity, "YYYY/MM/DD") +
                "' AND TODT >= '" +
                Format(Lmaturity, "YYYY/MM/DD") +
                "'";
              TRec.Open(mysql, Cnn, adOpenForwardOnly, adLockReadOnly);

              if (!TRec.EOF) {
                LSEBITax = Val(TRec.SEBITAX);
              }
            }

            MCLPattan = "A";
            LCBrokAmt = 0;

            if (LCloseRate !== 0) {
              LBrokQty = Math.abs(LClQty);

              if (LBrokType === "Z") LBrokQty = LBrokLot;

              if (LBrokType === "R") LBrokQty2 = LBrokLot;

              if (LBrokType === "A") LBrokQty2 = LBrokLot;

              if (LBrokType === "O") {
                LCBrokAmt = Calc_Brokerage(
                  LBrokType,
                  LBrokRate,
                  LBrokRate2,
                  0,
                  Math.abs(LClQty),
                  LCloseRate,
                  LCalval,
                  LInstType,
                  LStrike,
                  0,
                  LTradeableLot,
                  "B",
                  "2",
                  LItemID
                );
              } else if (LBrokType === "R" && LCloseRate === 0) {
                LCBrokAmt = 0;
              } else if (LBrokType === "A") {
                LCBrokAmt = 0;
              } else {
                LCBrokAmt = Calc_Brokerage(
                  LBrokType,
                  LBrokRate,
                  LBrokRate2,
                  LBrokQty,
                  Math.abs(LClQty),
                  LCloseRate,
                  LCalval,
                  LInstType,
                  LStrike,
                  0,
                  LTradeableLot,
                  "B",
                  "2",
                  LItemID
                );

                if (
                  LMinRate !== 0 &&
                  LMinRate * (LClQty / LRefLot) > LCBrokAmt
                ) {
                  LCBrokAmt = LMinRate * (LClQty / LRefLot);
                }
              }
            }

            LBrokAmount = LBrokAmount + LCBrokAmt;
            LSTTRate = 0;
            LSTTAmt =
              (LSTTRate / 100) * Math.abs(LClQty) * LCloseRate * LCalval;
            LTotSTTAmt = LTotSTTAmt + LSTTAmt;
            LcTranAmt = 0;

            if (LATranType === "T") {
              LcTranAmt = LTranRate * Math.abs(LClQty);
              LTranAmount = LTranAmount + LcTranAmt;
            } else if (LATranType === "P") {
              LcTranAmt =
                (LTranRate / 100) * Math.abs(LClQty) * LCloseRate * LCalval;
              LTranAmount = LTranAmount + LcTranAmt;
            }

            if (LSEBIType !== "N") {
              LSEBITaxAmt =
                (LSEBITax / 100) * (Math.abs(LClQty) * LCloseRate * LCalval);
            }

            LTotSEBITaxAmt = LTotSEBITaxAmt + LSEBITaxAmt;

            if (LCGSTType === "1")
              LCGSTAmt = LCGSTAmt + (LCBrokAmt * LCGSTRate) / 100;

            if (LSGSTType === "1")
              LSGSTAmt = LSGSTAmt + (LCBrokAmt * LSGSTRate) / 100;

            if (LIGSTType === "1")
              LIGSTAmt = LIGSTAmt + (LCBrokAmt * LIGSTRate) / 100;

            if (LUTTType === "1")
              LUTTAmt = LUTTAmt + (LCBrokAmt * LUTTRate) / 100;

            GSrvTax = GSrvTax + (LCBrokAmt * MSrvTax) / 100;
            LSBCTaxAmt = LSBCTaxAmt + (LCBrokAmt * LSBCTax) / 100;

            if (GSrStdYN === "Y") {
              if (Lmaturity >= DateValue(GSrvDate)) {
                GSrvTax = GSrvTax + (LcTranAmt * MSrvTax) / 100;
                LSBCTaxAmt = LSBCTaxAmt + (LcTranAmt * LSBCTax) / 100;

                if (LCGSTType === "1")
                  LCGSTAmt = LCGSTAmt + (LcTranAmt * LCGSTRate) / 100;

                if (LSGSTType === "1")
                  LSGSTAmt = LSGSTAmt + (LcTranAmt * LSGSTRate) / 100;

                if (LIGSTType === "1")
                  LIGSTAmt = LIGSTAmt + (LcTranAmt * LIGSTRate) / 100;

                if (LUTTType === "1")
                  LUTTAmt = LUTTAmt + (LcTranAmt * LUTTRate) / 100;
              }

              if (Lmaturity >= DateValue(GSrvDate2)) {
                LCGSTAmt = LCGSTAmt + (LSEBITaxAmt * LCGSTRate) / 100;
                LSGSTAmt = LSGSTAmt + (LSEBITaxAmt * LSGSTRate) / 100;
                LIGSTAmt = LIGSTAmt + (LSEBITaxAmt * LIGSTRate) / 100;
              }
            }
          }
          GSTDChrs = 0;
          LSTTRate = 0;
          LSTTAmt = 0;

          if (req?.body?.GRoundOff === "Y") {
            LTranAmount = Math.round(LTranAmount) * -1;
            LBrokAmount = Math.round(LBrokAmount) * -1;
            LTotSEBITaxAmt = Math.round(LTotSEBITaxAmt) * -1;
          } else {
            LTranAmount = LTranAmount * -1;
            LBrokAmount = LBrokAmount * -1;
            LTotSEBITaxAmt = LTotSEBITaxAmt * -1;
          }

          if (LRiskMType !== "N" && LRiskMApp === "Y") {
            LRiskMAmt = Calc_RiskMFees(
              LSaudaCode,
              LAExCode,
              LPartyCode.toString(),
              LTFromDate,
              LTillDate,
              Lmaturity,
              LCalval,
              LItemID,
              LSaudaID,
              LACCID
            );
          }

          LShareAmt = 0;

          if (req?.body?.GUniqClientId === "BRO2-CHE") {
            if (req?.body?.Check4 === 1) {
              LShareAmt = Calc_SharePAmt(
                LSaudaID,
                LPartyCode,
                LTFromDate,
                LTillDate
              );
            }
          } else {
            if (req?.body?.Check4 === 1) {
              LShareAmt = Calc_ShareAmt(
                LSaudaID,
                LPartyCode,
                LTFromDate,
                LTillDate
              );
            }
          }

          if (LRiskMAmt !== 0) {
            if (LRiskMType === "P") {
              LRiskMAmt = LRiskMAmt * -1;
            }

            GSrvTax = GSrvTax + (MSrvTax / 100) * (LRiskMAmt * -1);
            LSBCTaxAmt = LSBCTaxAmt + LRiskMAmt * -1 * (LSBCTax / 100);

            if (LCGSTType === "1") {
              LCGSTAmt = LCGSTAmt + (LRiskMAmt * -1 * LCGSTRate) / 100;
            }

            if (LSGSTType === "1") {
              LSGSTAmt = LSGSTAmt + (LRiskMAmt * -1 * LSGSTRate) / 100;
            }

            if (LIGSTType === "1") {
              LIGSTAmt = LIGSTAmt + (LRiskMAmt * -1 * LIGSTRate) / 100;
            }

            if (LUTTType === "1") {
              LUTTAmt = LUTTAmt + (LRiskMAmt * -1 * LUTTRate) / 100;
            }
          }

          LTotStmAmt = LTotStmAmt * -1;
          LTotSTTAmt = LTotSTTAmt * -1;
          LCGSTAmt = LCGSTAmt * -1;
          LSGSTAmt = LSGSTAmt * -1;
          LIGSTAmt = LIGSTAmt * -1;
          LUTTAmt = LUTTAmt * -1;

          if (LTotSEBITaxAmt === 0 && req?.body?.GUniqClientId === "BRO2-CHE") {
            if (Check4.Value === 1) {
              LTotSEBITaxAmt = Calc_SharebAmt(
                LSaudaID,
                LPartyCode,
                LTFromDate,
                LTillDate
              );
            }
          }
          if (LBrokType === "3") {
            LCloseRate = SDCLRATE(
              LSaudaID,
              Lmaturity <= LTillDate ? Lmaturity : LTillDate,
              "C"
            );
            LBrokAmount = Calculate_Percentage_BrokAmt(
              lOpQty,
              LTillDate,
              LACCID,
              LSaudaID,
              LTotOpAmt,
              LBrokRate,
              Lmaturity,
              LCloseRate
            );
            LBrokAmount = LBrokAmount * -1;
          }

          if (req?.body?.GStandingYN === "Y") {
            GSTDChrs = standingAmt(
              LSaudaCode,
              LAExCode,
              LPartyCode.toString(),
              LTFromDate,
              LTillDate,
              Lmaturity,
              LItemCode,
              LSaudaID
            );
          }
          if (req?.body?.GRoundOff === "Y") {
            GSTDChrs = Math.round(GSTDChrs);
          }
          GSrvTax = -1 * GSrvTax;
          LSBCTaxAmt = LSBCTaxAmt * -1;
          if (LNetPosition !== 0) {
            // Party Rate ujjwal
            if (LBillBy === "B") {
              if (LClQty > 0) {
                LCloseRate = LBuyClRate;
              } else {
                LCloseRate = LSellClRate;
              }
            }
            if (
              req?.body?.MFormat === "Bill Summary" ||
              req?.body?.MFormat === "Bill Summary With Sharing" ||
              req?.body?.MFormat === "Account Statement Summary"
            ) {
              LCloseAmt = Math.abs(parseFloat(LClQty)) * LCloseRate * LCalval;
              if (LClQty > 0) {
                LTotSellAmt = LTotSellAmt + LCloseAmt;
              } else {
                LTotBuyAmt = LTotBuyAmt + LCloseAmt;
              }
            }
            // If (LInstType === "CSH" && CshMTMchk.Value === 0) || (LInstType === "OPT" && OptMTMChk.Value === 0 && LMaturity < LTillDate)
            if (LInstType === "CSH" && CshMTMchk.Value === 0) {
              AccRecSet.Filter = 0;
              AccRecSet.Filter =
                LPartyCode.toString() + " AND SAUDACODE='" + LSaudaCode + "'";
              AccRecSet.Sort = "SRNO DESC";
              if (!AccRecSet.EOF) {
                AccRecSet.Fields("BROKERAGE").Value = parseFloat(LBrokAmount);
                AccRecSet.Fields("STANDING").Value = GSTDChrs;
                AccRecSet.Fields("TRFEES").Value = parseFloat(LTranAmount);
                if (!LStmFlag) {
                  AccRecSet.Fields("TurnOverTax").Value =
                    parseFloat(LTotStmAmt) + parseFloat(LNStmAmt);
                  LStmFlag = true;
                } else {
                  AccRecSet.Fields("TurnOverTax").Value =
                    parseFloat(LTotStmAmt);
                }
                AccRecSet.Fields("STTTax").Value = LTotSTTAmt;
                AccRecSet.Fields("TRANTAX").Value = parseFloat(LRiskMAmt);
                AccRecSet.Fields("SEBITAXAMT").Value = LTotSEBITaxAmt;
                AccRecSet.Fields("ServiceTax").Value = GSrvTax;
                AccRecSet.Fields("SBCTAXAMT").Value = LSBCTaxAmt;
                AccRecSet.Fields("PrevBal").Value = LprevBal;
                AccRecSet.Fields("DEBITAMT").Value = LNetRecdPay;
                AccRecSet.Fields("CREDITAMT").Value = LNetMargin;
                AccRecSet.Fields("SHAREAMT").Value = LShareAmt;
                if (req?.body?.Check12 === 0) {
                  AccRecSet.Fields("CGSTAMT").Value = LCGSTAmt;
                  AccRecSet.Fields("SGSTAMT").Value = LSGSTAmt;
                }
                AccRecSet.Fields("IGSTAMT").Value = LIGSTAmt;
                AccRecSet.Fields("UTTAMT").Value = LUTTAmt;
                AccRecSet.Update();
              } else {
                if (req?.body?.Check10 === 1) {
                  if (
                    lOpQty !== 0 ||
                    LClQty !== 0 ||
                    LTotSellQty ||
                    LTotBuyQty !== 0
                  ) {
                    AccRecSet.AddNew();
                    AccRecSet.Fields("BILLNO").Value = LSetNo;
                    AccRecSet.Fields("RPTGRP").Value = 1;
                    AccRecSet.Fields("SHAREAMT").Value = LShareAmt;
                    CountRec = CountRec + 1;
                    AccRecSet.Fields("SRNO").Value = CountRec;
                    AccRecSet.Fields("SAUDACODE").Value = LSaudaCode;
                    AccRecSet.Fields("EXCODE").Value = LAExCode;
                    AccRecSet.Fields("ITEMCODE").Value = LItemCode;
                    AccRecSet.Fields("DRAMOUNT").Value = Math.abs(
                      parseFloat(LTotBuyAmt)
                    );
                    AccRecSet.Fields("TRFEES").Value = parseFloat(LTranAmount);
                    AccRecSet.Fields("CRAMOUNT").Value = Math.abs(
                      parseFloat(LTotSellAmt)
                    );
                    AccRecSet.Fields("CTYPE").Value = "B";
                    AccRecSet.Fields("PATTAN").Value = MCLPattan;
                    AccRecSet.Fields("BROKERAGE").Value = LBrokAmount;
                    AccRecSet.Fields("REFLOT").Value = LRefLot;
                    AccRecSet.Fields("STANDING").Value = GSTDChrs;
                    AccRecSet.Fields("BRATE").Value = LCloseRate;
                    AccRecSet.Fields("SRATE").Value = LCloseRate;
                    AccRecSet.Fields("BQNTY").Value = Math.abs(
                      parseFloat(LTotBuyQty)
                    );
                    AccRecSet.Fields("SQNTY").Value = Math.abs(
                      parseFloat(LTotSellQty)
                    );
                    AccRecSet.Fields("DEBITAMT").Value = Math.abs(
                      parseFloat(LTotBQAmt)
                    );
                    AccRecSet.Fields("CREDITAMT").Value = Math.abs(
                      parseFloat(LTotSQAmt)
                    );
                    AccRecSet.Fields("OPENQTY").Value = parseFloat(lOpQty);
                    AccRecSet.Fields("OPENRATE").Value = parseFloat(LOpenRate);
                    Add_PartyDetails();
                    if (!LStmFlag) {
                      AccRecSet.Fields("TurnOverTax").Value =
                        parseFloat(LTotStmAmt) + parseFloat(LNStmAmt);
                      LStmFlag = true;
                    } else {
                      AccRecSet.Fields("TurnOverTax").Value =
                        parseFloat(LTotStmAmt);
                    }
                    AccRecSet.Fields("STTTAX").Value = LTotSTTAmt;
                    AccRecSet.Fields("TRANTAX").Value = parseFloat(LRiskMAmt);
                    AccRecSet.Fields("SEBITAXAMT").Value = LTotSEBITaxAmt;
                    AccRecSet.Fields("CONTDATE1").Value = Format(
                      LTillDate,
                      "yyyy/MM/dd"
                    );
                    AccRecSet.Fields("CALVAL").Value = LCalval;
                    AccRecSet.Fields("INVDATE").Value = Format(
                      Lmaturity,
                      "yyyy/MM/dd"
                    );
                    AccRecSet.Fields("ServiceTax").Value = GSrvTax;
                    AccRecSet.Fields("SBCTAXAMT").Value = LSBCTaxAmt;
                    AccRecSet.Fields("PrevBal").Value = LprevBal;
                    AccRecSet.Fields("BrokRate").Value = LBrokRate;
                    AccRecSet.Fields("MarRate").Value = LMarRate;
                    AccRecSet.Fields("BrokType").Value = LBrokType;
                    if (req?.body?.Check12 === 0) {
                      AccRecSet.Fields("CGSTAMT").Value = LCGSTAmt;
                      AccRecSet.Fields("SGSTAMT").Value = LSGSTAmt;
                    }
                    AccRecSet.Fields("IGSTAMT").Value = LIGSTAmt;
                    AccRecSet.Fields("UTTAMT").Value = LUTTAmt;
                    AccRecSet.Fields("CONTIME").Value = vbNullString;
                    AccRecSet.Fields("CONNO").Value = vbNullString;
                    AccRecSet.Fields("MarAmt").Value = LMarAmt;
                    AccRecSet.Fields("MarType").Value = LMarType;
                    AccRecSet.Fields("rpattan").Value = vbNullString;
                    AccRecSet.Fields("SBROKRATE").Value = 0;
                    AccRecSet.Fields("SBROKType").Value = vbNullString;
                    AccRecSet.Fields("INVNO").Value = 0;
                    AccRecSet.Fields("CONTDATE").Value = vbNullString;
                    AccRecSet.Update();
                  }
                }
              }
            } else {
              var mstanding = false;
              mstanding = false;
              if (LTillDate >= Lmaturity && LClQty !== 0 && LBrokType === "3") {
                MCLPattan = "A";
              }
              if (
                Math.abs(parseFloat(LTotBuyQty)) +
                  Math.abs(parseFloat(LTotSellQty)) +
                  Math.abs(parseFloat(LClQty)) +
                  Math.abs(parseFloat(lOpQty)) !==
                0
              ) {
                if (req?.body?.Check12 === 0) {
                  AccRecSet.AddNew();
                  mstanding = true;
                  CountRec = CountRec + 1;
                  AccRecSet.Fields("SRNO").Value = CountRec;
                  AccRecSet.Fields("BILLNO").Value = LSetNo;
                  AccRecSet.Fields("SAUDACODE").Value = LSaudaCode;
                  if (req?.body?.Check10 === 1) {
                    AccRecSet.Fields("RPTGRP").Value = 2;
                  } else {
                    AccRecSet.Fields("RPTGRP").Value = 1;
                  }
                  AccRecSet.Fields("SHAREAMT").Value = parseFloat(LShareAmt);
                  AccRecSet.Fields("REFLOT").Value = LRefLot;
                  AccRecSet.Fields("EXCODE").Value = LAExCode;
                  AccRecSet.Fields("ITEMCODE").Value = LItemCode;
                  AccRecSet.Fields("CREDITAMT").Value = parseFloat(LNetMargin);
                  AccRecSet.Fields("ServiceTax").Value = GSrvTax;
                  AccRecSet.Fields("SBCTAXAMT").Value = LSBCTaxAmt;
                  AccRecSet.Fields("STTTax").Value = LTotSTTAmt;
                  AccRecSet.Fields("SEBITAXAMT").Value = LTotSEBITaxAmt;
                  AccRecSet.Fields("TRANTAX").Value = parseFloat(LRiskMAmt);
                  AccRecSet.Fields("TRFEES").Value = parseFloat(LTranAmount);
                  AccRecSet.Fields("INVDATE").Value = Format(
                    Lmaturity,
                    "yyyy/MM/dd"
                  );
                  AccRecSet.Fields("CALVAL").Value = LCalval;
                  AccRecSet.Fields("PrevBal").Value = LprevBal;
                  AccRecSet.Fields("MarRate").Value = LMarRate;
                  AccRecSet.Fields("MarAmt").Value = LMarAmt;
                  AccRecSet.Fields("MarType").Value = LMarType;
                  AccRecSet.Fields("SBROKRATE").Value = LBrokRate;
                  AccRecSet.Fields("BrokType").Value = LBrokType;
                  AccRecSet.Fields("SBROKType").Value = LBrokType;
                  if (req?.body?.Check12 === 0) {
                    AccRecSet.Fields("CGSTAMT").Value = parseFloat(LCGSTAmt);
                    AccRecSet.Fields("SGSTAMT").Value = parseFloat(LSGSTAmt);
                  }
                  AccRecSet.Fields("IGSTAMT").Value = parseFloat(LIGSTAmt);
                  AccRecSet.Fields("UTTAMT").Value = parseFloat(LUTTAmt);
                  AccRecSet.Fields("PATTAN").Value = MCLPattan;
                  AccRecSet.Fields("OPENQTY").Value = parseFloat(lOpQty);
                  AccRecSet.Fields("OPENRATE").Value = parseFloat(LOpenRate);
                  Add_PartyDetails();
                  if (parseFloat(LClQty) > 0) {
                    AccRecSet.Fields("CONTDATE").Value = Format(
                      LTillDate,
                      "yyyy/MM/dd"
                    );
                    AccRecSet.Fields("SQNTY").Value = Math.abs(
                      parseFloat(LClQty)
                    );
                    AccRecSet.Fields("SRATE").Value = LCloseRate;
                    AccRecSet.Fields("CTYPE").Value = "S";
                    if (MFormat === "Account Statement") {
                      AccRecSet.Fields("CRAMOUNT").Value =
                        Math.abs(parseFloat(LClQty)) * LCloseRate;
                    } else {
                      AccRecSet.Fields("DRAMOUNT").Value = Math.abs(
                        parseFloat(LTotBuyAmt)
                      );
                      AccRecSet.Fields("CRAMOUNT").Value = Math.abs(
                        parseFloat(LTotSellAmt)
                      );
                      AccRecSet.Fields("BQNTY").Value = Math.abs(
                        parseFloat(LTotBuyQty)
                      );
                      AccRecSet.Fields("SQNTY").Value = Math.abs(
                        parseFloat(LTotSellQty)
                      );
                      AccRecSet.Fields("DEBITAMT").Value = Math.abs(
                        parseFloat(LTotBQAmt)
                      );
                      AccRecSet.Fields("CREDITAMT").Value = Math.abs(
                        parseFloat(LTotSQAmt)
                      );
                    }
                    AccRecSet.Fields("BBROKAMT").Value = LCBrokAmt;
                    AccRecSet.Fields("SBROKAMT").Value = 0;
                  } else if (parseFloat(LClQty) < 0) {
                    AccRecSet.Fields("CONTDATE1").Value = Format(
                      LTillDate,
                      "yyyy/MM/dd"
                    );
                    AccRecSet.Fields("BQNTY").Value = Math.abs(
                      parseFloat(LClQty)
                    );
                    AccRecSet.Fields("BRATE").Value = LCloseRate;
                    AccRecSet.Fields("CTYPE").Value = "B";
                    AccRecSet.Fields("BBROKAMT").Value = LCBrokAmt;
                    if (MFormat === "Account Statement") {
                      AccRecSet.Fields("DRAMOUNT").Value =
                        Math.abs(parseFloat(LClQty)) * LCloseRate;
                    } else {
                      AccRecSet.Fields("DRAMOUNT").Value = Math.abs(
                        parseFloat(LTotBuyAmt)
                      );
                      AccRecSet.Fields("CRAMOUNT").Value = Math.abs(
                        parseFloat(LTotSellAmt)
                      );
                      AccRecSet.Fields("BQNTY").Value = Math.abs(
                        parseFloat(LTotBuyQty)
                      );
                      AccRecSet.Fields("SQNTY").Value = Math.abs(
                        parseFloat(LTotSellQty)
                      );
                      AccRecSet.Fields("DEBITAMT").Value = Math.abs(
                        parseFloat(LTotBQAmt)
                      );
                      AccRecSet.Fields("CREDITAMT").Value = Math.abs(
                        parseFloat(LTotSQAmt)
                      );
                    }
                    AccRecSet.Fields("BBROKAMT").Value = 0;
                    AccRecSet.Fields("SBROKAMT").Value = LCBrokAmt;
                  } else if (parseFloat(LClQty) === 0) {
                    AccRecSet.Fields("CONTDATE1").Value = Format(
                      LTillDate,
                      "yyyy/MM/dd"
                    );
                    AccRecSet.Fields("BQNTY").Value = Math.abs(
                      parseFloat(LClQty)
                    );
                    AccRecSet.Fields("BRATE").Value = LCloseRate;
                    AccRecSet.Fields("CTYPE").Value = "B";
                    AccRecSet.Fields("BBROKAMT").Value = LCBrokAmt;
                    if (MFormat === "Account Statement") {
                      AccRecSet.Fields("DRAMOUNT").Value =
                        Math.abs(parseFloat(LClQty)) * LCloseRate;
                    } else {
                      AccRecSet.Fields("DRAMOUNT").Value = Math.abs(
                        parseFloat(LTotBuyAmt)
                      );
                      AccRecSet.Fields("CRAMOUNT").Value = Math.abs(
                        parseFloat(LTotSellAmt)
                      );
                      AccRecSet.Fields("BQNTY").Value = Math.abs(
                        parseFloat(LTotBuyQty)
                      );
                      AccRecSet.Fields("SQNTY").Value = Math.abs(
                        parseFloat(LTotSellQty)
                      );
                      AccRecSet.Fields("DEBITAMT").Value = Math.abs(
                        parseFloat(LTotBQAmt)
                      );
                      AccRecSet.Fields("CREDITAMT").Value = Math.abs(
                        parseFloat(LTotSQAmt)
                      );
                    }
                    AccRecSet.Fields("BBROKAMT").Value = 0;
                    AccRecSet.Fields("SBROKAMT").Value = LCBrokAmt;
                  }
                  AccRecSet.Fields("BROKERAGE").Value = LBrokAmount;
                  AccRecSet.Fields("STANDING").Value = GSTDChrs;
                  if (LStmFlag === false) {
                    AccRecSet.Fields("TurnOverTax").Value =
                      parseFloat(LTotStmAmt) + parseFloat(LNStmAmt);
                    LStmFlag = true;
                  } else {
                    AccRecSet.Fields("TurnOverTax").Value =
                      parseFloat(LTotStmAmt);
                  }
                  AccRecSet.Fields("BINSTYPE").Value = "CSH";
                  AccRecSet.Update();
                }
              }
            }
          } else {
            AccRecSet.Filter = 0;
            AccRecSet.Filter =
              "'" + LPartyCode + "' AND SAUDACODE='" + LSaudaCode + "'";
            AccRecSet.Sort = "SRNO DESC";
            if (!AccRecSet.EOF) {
              AccRecSet.Fields("BROKERAGE").Value = LBrokAmount;
              AccRecSet.Fields("STANDING").Value = GSTDChrs;
              AccRecSet.Fields("TRFEES").Value = LTranAmount;
              if (LStmFlag === false) {
                AccRecSet.Fields("TurnOverTax").Value =
                  parseFloat(LTotStmAmt) + parseFloat(LNStmAmt);
                LStmFlag = true;
              } else {
                AccRecSet.Fields("TurnOverTax").Value = parseFloat(LTotStmAmt);
              }
              AccRecSet.Fields("SEBITAXAMT").Value = LTotSEBITaxAmt;
              AccRecSet.Fields("STTTax").Value = LTotSTTAmt;
              AccRecSet.Fields("TRANTAX").Value = LRiskMAmt;
              AccRecSet.Fields("rpattan").Value = "C";
              AccRecSet.Fields("ServiceTax").Value = GSrvTax;
              AccRecSet.Fields("SBCTAXAMT").Value = LSBCTaxAmt;
              AccRecSet.Fields("PrevBal").Value = LprevBal;
              AccRecSet.Fields("DEBITAMT").Value = LNetRecdPay;
              AccRecSet.Fields("CREDITAMT").Value = LNetMargin;
              AccRecSet.Fields("SHAREAMT").Value = LShareAmt;
              if (req?.body?.Check12 === 0) {
                AccRecSet.Fields("CGSTAMT").Value = LCGSTAmt;
                AccRecSet.Fields("SGSTAMT").Value = LSGSTAmt;
              }
              AccRecSet.Fields("IGSTAMT").Value = LIGSTAmt;
              AccRecSet.Fields("UTTAMT").Value = LUTTAmt;
              AccRecSet.Update();
            }
          }
          if (req?.body?.Check10 === 1) {
            TRec = null;
          }
          // >>> account statement report grouping
          if (
            req?.body?.Check10 === 1 &&
            Option5.Value &&
            req?.body?.Check12 === 0
          ) {
            // If LedgerFlag = True Then ''>>> account statement report grouping
            if (LedgerFlag === false) {
              TRec = null;
              TRec = new ADODB.Recordset();

              mysql =
                "SELECT A.VOU_NO,A.VOU_DT,a.AMOUNT, A.DR_CR, A.NARRATION FROM VCHAMT As A WHERE A.COMPCODE=" +
                req?.body?.companyCode +
                " And A.AC_CODE        '" +
                LPartyCode +
                "'";
              mysql +=
                " And VOU_TYPE <> 'S' AND A.VOU_DT >='" +
                Format(req?.body?.fromDate, "YYYY/MM/DD") +
                "' AND A.VOU_DT <='" +
                Format(vcDTP2.Value, "YYYY/MM/DD") +
                "'";
              mysql += "  ORDER BY A.VOU_DT,A.VOU_NO";

              TRec.Open(mysql, Cnn, adOpenStatic, adLockReadOnly);
              LBal = Get_ClosingBal(LACCID, req?.body?.fromDate - 1);
              if (LBal !== 0) {
                AccRecSet.AddNew();
                CountRec += 1;
                AccRecSet.Fields("RPTGRP").Value = 6;
                AccRecSet.Fields("SRNO").Value = CountRec;
                AccRecSet.Fields("CONNO").Value = "Opening Bal";
                AccRecSet.Fields("CONTDATE").Value = Format(
                  req?.body?.fromDate,
                  "yyyy/MM/dd"
                );
                AccRecSet.Fields("WORDAMT").Value = vbNullString;
                // buy sell
                if (LBal < 0) {
                  AccRecSet.Fields("DEBITAMT").Value = Math.abs(LBal);
                  AccRecSet.Fields("CREDITAMT").Value = 0;
                  AccRecSet.Fields("CTYPE").Value = "D";
                } else {
                  AccRecSet.Fields("DEBITAMT").Value = 0;
                  AccRecSet.Fields("CREDITAMT").Value = Math.abs(LBal);
                  AccRecSet.Fields("CTYPE").Value = "C";
                }
                AccRecSet.Fields("STANDING").Value = LBal;
                Add_PartyDetails();
                Add_OtherDetails();
                AccRecSet.Update();
              }
              while (!TRec.EOF) {
                AccRecSet.AddNew();
                CountRec += 1;
                AccRecSet.Fields("RPTGRP").Value = 6;
                AccRecSet.Fields("SRNO").Value = CountRec;
                AccRecSet.Fields("CONNO").Value = TRec.Fields("VOU_NO").Value;
                AccRecSet.Fields("CONTDATE").Value = Format(
                  TRec.Fields("VOU_DT").Value,
                  "yyyy/MM/dd"
                );
                AccRecSet.Fields("WORDAMT").Value =
                  TRec.Fields("NARRATION").Value;
                // buy sell
                if (TRec.Fields("DR_CR").Value === "D") {
                  AccRecSet.Fields("DEBITAMT").Value =
                    TRec.Fields("AMOUNT").Value;
                  AccRecSet.Fields("CREDITAMT").Value = 0;
                  AccRecSet.Fields("CTYPE").Value = "D";
                  LBal += TRec.Fields("AMOUNT").Value * -1;
                } else {
                  AccRecSet.Fields("DEBITAMT").Value = 0;
                  AccRecSet.Fields("CREDITAMT").Value =
                    TRec.Fields("AMOUNT").Value;
                  AccRecSet.Fields("CTYPE").Value = "C";
                  LBal += TRec.Fields("AMOUNT").Value;
                }
                AccRecSet.Fields("STANDING").Value = LBal;
                Add_PartyDetails();
                Add_OtherDetails();
                AccRecSet.Update();
                TRec.MoveNext();
              }
              LedgerFlag = true;
            }
          }
          CFlag = false;
          if (req?.body?.Check9 === 1) {
            if (req?.body?.Check12 === 0) {
              if (
                Math.abs(parseFloat(LTotBuyQty)) +
                  Math.abs(LTotSellQty) +
                  Math.abs(lOpQty) +
                  Math.abs(LClQty) !==
                0
              ) {
                if (
                  (req?.body?.Check10 === 1 && LCloseRate > 0) ||
                  Check9.Value === 1
                ) {
                  AccRecSet.AddNew();
                  CountRec += 1;
                  AccRecSet.Fields("BILLNO").Value = LSetNo;
                  AccRecSet.Fields("EXCODE").Value = LAExCode;

                  // old
                  if (req?.body?.Check10 === 1) {
                    AccRecSet.Fields("RPTGRP").Value = 4;
                  } else {
                    AccRecSet.Fields("RPTGRP").Value = 3;
                  }

                  // >>> account statement report grouping
                  if (Check9.Value === 1) {
                    AccRecSet.Fields("RPTGRP").Value = 4; // >>> summary
                  } else {
                    AccRecSet.Fields("RPTGRP").Value = 3; // >>> standing > net position
                  }

                  AccRecSet.Fields("SRNO").Value = CountRec;
                  AccRecSet.Fields("SAUDACODE").Value = LSaudaCode;
                  AccRecSet.Fields("CONTDATE").Value = Format(
                    LTillDate,
                    "yyyy/MM/dd"
                  );
                  AccRecSet.Fields("ITEMCODE").Value = LItemCode;
                  AccRecSet.Fields("REFLOT").Value = LRefLot;
                  AccRecSet.Fields("OPENQTY").Value = parseFloat(lOpQty);
                  AccRecSet.Fields("CLOSEQTY").Value = parseFloat(LClQty);
                  AccRecSet.Fields("PrevBal").Value = LprevBal;
                  AccRecSet.Fields("DEBITAMT").Value = LNetRecdPay;
                  AccRecSet.Fields("CREDITAMT").Value = LNetMargin;
                  if (LTotOpAmt > 0) {
                    LTotBuyAmt += LTotOpAmt;
                  } else {
                    LTotSellAmt += Math.abs(LTotOpAmt);
                  }
                  if (LClQty > 0) {
                    LTotSellAmt +=
                      Math.abs(parseFloat(LClQty)) * LCloseRate * LCalval;
                  } else {
                    LTotBuyAmt +=
                      Math.abs(parseFloat(LClQty)) * LCloseRate * LCalval;
                  }
                  AccRecSet.Fields("GROSSMTM").Value = LTotSellAmt - LTotBuyAmt;
                  AccRecSet.Fields("BROKAMT").Value = LBrokAmount;
                  AccRecSet.Fields("NETAMT").Value =
                    LTotSellAmt - LTotBuyAmt + LBrokAmount;

                  Add_PartyDetails();

                  // AccRecSet.Fields("BQNTY") = Abs(Val(LTotBuyQty)); AccRecSet.Fields("SQNTY") = LTotSellQty;
                  AccRecSet.Fields("SRATE").Value = LCloseRate;
                  AccRecSet.Fields("CTYPE").Value = "S";
                  AccRecSet.Fields("PATTAN").Value = MCLPattan;
                  AccRecSet.Fields("INVDATE").Value = Format(
                    Lmaturity,
                    "yyyy/MM/dd"
                  );
                  AccRecSet.Fields("CALVAL").Value = LCalval;

                  AccRecSet.Update();
                }
              }
            }
          }
          // END CLOSING
          // Update Interest
          updinterest(LPartyCode, LSaudaCode_muliple);
          // DoEvents - This function is specific to Visual Basic for Applications (VBA) and is used to allow the operating system to process other events in the queue.
          // There is no direct equivalent in JavaScript as it's not event-driven in the same way. You may not need to replicate this behavior depending on the context.
          if (partiesRecordset.EOF) {
            flag_party(LPartyCode, LPartyName);
          }
        }
        if (TempPart !== LPartyCode) {
          TempPart = LPartyCode;
          if (
            req?.body?.Check10 === 1 &&
            Option5.Value &&
            req?.body?.Check12 === 1 &&
            MFormat === "Account Statement"
          ) {
            if (LedgerFlag === false) {
              TRecVOU = null;
              TRecVOU = new ADODB.Recordset();
              let mysql =
                "SELECT A.VOU_NO,A.VOU_DT,a.AMOUNT, A.DR_CR, A.NARRATION FROM VCHAMT As A WHERE A.COMPCODE=" +
                req?.body?.companyCode +
                " And A.AC_CODE '" +
                LPartyCode +
                "'";
              if (Check11.Value === 1) {
                mysql +=
                  " And VOU_TYPE <> 'S' AND A.VOU_DT ='" +
                  Format(loopdate, "YYYY/MM/DD") +
                  "'";
              } else {
                mysql +=
                  " And VOU_TYPE <> 'S' AND VOU_TYPE <> 'I' AND A.VOU_DT >='" +
                  Format(req?.body?.fromDate, "YYYY/MM/DD") +
                  "' and A.VOU_DT <='" +
                  Format(vcDTP2.Value, "YYYY/MM/DD") +
                  "' ";
              }
              mysql += " ORDER BY A.VOU_DT,A.VOU_NO";
              TRecVOU.Open(mysql, Cnn, adOpenStatic, adLockReadOnly);
              while (!TRecVOU.EOF) {
                CountRec = CountRec + 1;
                Add_VOUCHER(
                  "2",
                  CountRec,
                  "TRAN",
                  TRecVOU.VOU_DT,
                  TRecVOU.VOU_NO,
                  TRecVOU.NARRATION,
                  TRecVOU.DR_CR,
                  TRecVOU.AMOUNT,
                  0
                );
                TRecVOU.MoveNext();
              }
            }
          }
          CountRec = CountRec + 1;
        }
      }
      Client_Margin();
    } catch (error) {
      console.error("Error:", error);
      res.status(500).json({ message: "Internal Server Error" });
    }
  },
};

module.exports = testController;
