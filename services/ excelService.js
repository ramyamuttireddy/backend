const XLSX = require("xlsx");
const credential = require("../config/azure");

async function getExcelData() {
  try {
    const tokenResponse = await credential.getToken(
      "https://graph.microsoft.com/.default"
    );

    const accessToken = tokenResponse.token;

    const url =
      `https://graph.microsoft.com/v1.0/drives/${process.env.DRIVE_ID}` +
      `/items/${process.env.FILE_ID}/content`;

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(errorText);
    }

    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const workbook = XLSX.read(buffer, {
      type: "buffer",
    });

    const sheetNames = workbook.SheetNames;

    console.log("Sheets:", sheetNames);

    const firstSheetName = sheetNames[0];

    const worksheet =
      workbook.Sheets[firstSheetName];

    const data =
      XLSX.utils.sheet_to_json(worksheet);

    return {
      sheetName: firstSheetName,
      data,
    };
  } catch (error) {
    console.error(
      "Excel Service Error:",
      error
    );

    throw error;
  }
}

module.exports = {
  getExcelData,
};