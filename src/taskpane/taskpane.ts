import { CompanyFullInfo } from "../types";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values, address");
      await context.sync();

      // Extract IDs from the selected range
      const ids = range.values.map((row) => row[2]); // Assuming the ID is in the third column (index 2)

      // Fetch data for the extracted IDs
      const data = await fetchData(ids);

      // Populate Excel with the fetched data
      await populateExcel(data, range.address);
    });
  } catch (error) {
    console.error(error);
  }
}

async function fetchData(ids: string[]): Promise<CompanyFullInfo[]> {
  try {
    const results: CompanyFullInfo[] = [];

    for (const id of ids) {
      let response = await fetch(`https://apiba.prgapp.kz/CompanyFullInfo?id=${id}&lang=ru`, {
        headers: {
          accept: "*/*",
          "accept-language": "en-GB,en-US;q=0.9,en;q=0.8,ru;q=0.7",
          priority: "u=1, i",
          "sec-ch-ua": '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
          "sec-ch-ua-mobile": "?0",
          "sec-ch-ua-platform": '"Windows"',
          "sec-fetch-dest": "empty",
          "sec-fetch-mode": "cors",
          "sec-fetch-site": "cross-site",
          Referer: "https://ba.prg.kz/",
          "Referrer-Policy": "strict-origin-when-cross-origin",
        },
        body: null,
        method: "GET",
      });

      let data: CompanyFullInfo = await response.json();
      results.push(data);
    }

    return results;
  } catch (error) {
    console.error("Error fetching data:", error);
    return [];
  }
}

async function populateExcel(data: CompanyFullInfo[], selectedRangeAddress: string) {
  console.log("selectedRangeAddress", selectedRangeAddress);

  // Map the data to the desired columns
  const mappedData = data.map((result) => [
    result.basicInfo.ceo.value.title,
    result.basicInfo.addressRu.value,
    result.gosZakupContacts?.phone.length > 0 ? result.gosZakupContacts.phone.map((item) => item.value).join("; ") : "",
    result.basicInfo.registrationDate.value,
    result.basicInfo.primaryOKED.value,
    result.basicInfo.secondaryOKED.value.join("; "),
  ]);

  // Headers
  const headers = ["name", "address", "phone", "registration", "primary", "secondary"];

  // Populate the Excel sheet
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const selectedRange = sheet.getRange(selectedRangeAddress);
    selectedRange.load("columnIndex, rowIndex, rowCount");
    await context.sync();

    // Calculate the starting cell for the new data
    const startColumn = selectedRange.columnIndex + 6; // Assuming the Date column is the 6th column (index 5)
    const startRow = selectedRange.rowIndex;

    // Insert headers
    const headerRange = sheet.getRangeByIndexes(0, startColumn, 1, headers.length);
    headerRange.values = [headers];

    // Insert data
    const dataRange = sheet.getRangeByIndexes(startRow, startColumn, mappedData.length, mappedData[0].length);
    dataRange.values = mappedData;

    await context.sync();
  });
}
