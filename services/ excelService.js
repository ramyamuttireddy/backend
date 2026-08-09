const XLSX = require("xlsx");
const credential = require("../config/azure");

async function getExcelData() {
  try {
    // 1. Get Microsoft Graph access token
    const tokenResponse = await credential.getToken(
      "https://graph.microsoft.com/.default"
    );

    const accessToken = tokenResponse.token;

    // 2. Excel file URL
    const url =
      `https://graph.microsoft.com/v1.0/drives/${process.env.DRIVE_ID}` +
      `/items/${process.env.FILE_ID}/content`;

    // 3. Download Excel file
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (!response.ok) {
      const errorText = await response.text();

      throw new Error(errorText);
    }

    // 4. Convert response to buffer
    const arrayBuffer = await response.arrayBuffer();

    const buffer = Buffer.from(arrayBuffer);

    // 5. Read Excel workbook
    const workbook = XLSX.read(buffer, {
      type: "buffer",
    });

    // 6. Get sheet names
    const sheetNames = workbook.SheetNames;

    console.log("Sheets:", sheetNames);

    // 7. First sheet
    const firstSheetName = sheetNames[0];

    const worksheet = workbook.Sheets[firstSheetName];

    // 8. Convert Excel sheet → JSON
    const data = XLSX.utils.sheet_to_json(worksheet);

    return {
      sheetName: firstSheetName,
      data,
    };
  } catch (error) {
    console.error("Excel Service Error:", error);

    throw error;
  }
}

module.exports = {
  getExcelData,
};