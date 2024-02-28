const { adobeConnect } = require("../utils/utils");

class partyModel {
  static async getPartiesBySelection(companyCode, toDate) {
    try {
      const query = `
        SELECT DISTINCT AC_CODE, NAME 
        FROM ACCOUNTD 
        WHERE COMPCODE = ${companyCode}
        AND AC_CODE IN (
          SELECT DISTINCT PARTY 
          FROM CTR_D 
          WHERE COMPCODE = ${companyCode}
          AND Condate <= '${toDate}'
          AND SAUDA IN (
            SELECT DISTINCT SAUDACODE 
            FROM SAUDAMAST 
            WHERE COMPCODE = ${companyCode}
            AND MATURITY >= '${toDate}'
          )
        ) 
        ORDER BY NAME`;

      const result = await adobeConnect(query);
      return result;
    } catch (error) {
      console.error("Error:", error);
      throw error; // Rethrow the error to propagate it to the caller
    }
  }

  static async getPartyRecordsBySelection(req) {
    try {
      const query = `
              SELECT DISTINCT A.ACCID,A.AC_CODE,A.NAME,A.OP_BAL,
              B.ADD1,B.CITY, B.PANNO, B.PARTYTYPE, B.PIN, B.PHONEO, B.PHONER, B.FAX, B.MOBILE, B.PIN, B.SRTAXAPP, B.CGST, B.SGST, B.IGST, B.UTT, B.CTTTYPE,B.RISKMTYPE,B.OPTCUTBROK,B.FUTCUTBROK,
              B.APPLYON,B.EMAIL,B.GSTIN,B.STATE,B.STATECODE,B.SEBITYPE
              FROM ACCOUNTM AS A, ACCOUNTD AS B, CTR_D AS C
              WHERE A.COMPCODE = ${req?.body?.companyCode} AND A.ACCID = B.ACCID
               AND C.CONDATE <= '${req?.body?.toDate}'
              AND A.ACCID = C.ACCID 
              ${
                req?.body?.AllFmly === false &&
                req?.body?.LSFmlyIDs.length > 0 &&
                AllParties === true
                  ? `AND A.ACCID IN (SELECT ACCID FROM ACCFMLYD WHERE FMLYID IN (${req?.body?.LSFmlyIDs}))`
                  : ""
              }
              ${
                req?.body?.allParties === false &&
                req?.body?.selectedParties.length > 0
                  ? `AND A.AC_CODE IN (${req?.body?.selectedParties})`
                  : ""
              }
              ORDER BY A.NAME
          `;

      const result = await adobeConnect(query);
      return result;
    } catch (error) {
      console.error("Error:", error);
      throw error; // Rethrow the error to propagate it to the caller
    }
  }

  static async getTradedItemDetailsBySelection(req) {
    try {
      const query = `
              SELECT DISTINCT 
             
              S.SAUDACODE,S.SAUDAID,S.INSTTYPE,I.ITEMCODE,S.STRIKEPRICE,S.OPTTYPE,S.MATURITY,S.TRADEABLELOT,S.BROKLOT ,S.REFLOT,
              I.ITEMID,I.LOT,I.RISKMAPP,I.SCGROUP,EX.EXCODE,EX.EXID,EX.LOTWISE,EX.CONTRACTACC
              FROM  CTR_D AS C, ITEMMAST AS I, SAUDAMAST AS S ,EXMAST EX
              WHERE S.ITEMID = I.ITEMID AND S.MATURITY >= '${
                req?.body?.fromDate
              }' AND C.CONDATE <= '${req?.body?.toDate}' 
			        AND EX.EXID = S.EXID 
              ${
                req?.body?.allExchangeCodes === false &&
                req?.body?.selectedExchangeCodes.length > 0
                  ? `AND I.EXID IN (${req?.body?.selectedExchangeCodes})`
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
              ORDER BY EX.EXCODE,I.ITEMCODE,S.INSTTYPE,S.MATURITY
          `;

      const result = await adobeConnect(query);
      return result;
    } catch (error) {
      console.error("Error:", error);
      throw error; // Rethrow the error to propagate it to the caller
    }
  }
}

module.exports = partyModel;
