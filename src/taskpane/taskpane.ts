import { CompanyFullInfo } from "../types";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const appBody = document.getElementById("app-body");
    if (appBody) {
      appBody.style.display = "flex";
    }

    const runButton = document.getElementById("run");
    if (runButton) {
      runButton.onclick = run;
    }
  }
});

export async function run() {
  const runButton = document.getElementById("run") as HTMLButtonElement;
  runButton.disabled = true; // Disable the button
  runButton.classList.add("disabled"); // Add a class for styling

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("rowIndex, rowCount");
      await context.sync();

      // Calculate the new range address (columns A to F)
      const startRow = range.rowIndex + 1; // Excel row indices are 1-based
      const endRow = startRow + range.rowCount - 1;
      const newRangeAddress = `A${startRow}:F${endRow}`;

      // Get the new range that spans columns A to F
      const newRange = context.workbook.worksheets.getActiveWorksheet().getRange(newRangeAddress);
      newRange.load("values");
      await context.sync();

      // Extract IDs from the new range
      const ids = newRange.values.map((row) => row[2]); // Assuming the ID is in the third column (index 2)

      // Fetch data for the extracted IDs
      const data = await fetchData(ids);

      // Populate Excel with the fetched data
      await populateExcel(data, newRangeAddress);
    });
  } catch (error) {
    console.error(error);
  } finally {
    runButton.disabled = false; // Re-enable the button
    runButton.classList.remove("disabled"); // Remove the class for styling
  }
}

async function fetchData(ids: string[]): Promise<CompanyFullInfo[]> {
  try {
    const results: CompanyFullInfo[] = [];
    const defaultBasicInfo = {
      registrationDate: { value: null },
      onMarket: null,
      ceo: { value: { title: "Пользователь не найден" } },
      primaryOKED: { value: "" },
      secondaryOKED: { value: null },
      addressRu: { value: "" },
    };

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

      if (!response.ok) {
        results.push({ basicInfo: defaultBasicInfo, gosZakupContacts: null });
        continue;
      }

      const data: CompanyFullInfo = await response.json();
      if (data) results.push(data);
    }

    return results;
  } catch (error) {
    console.error("Error fetching data:", error);
    return [];
  }
}

async function populateExcel(data: CompanyFullInfo[], selectedRangeAddress: string) {
  // Map the data to the desired columns
  const mappedData = data.map((result) => [
    result.basicInfo.ceo.value?.title,
    result.basicInfo.addressRu.value,
    result.gosZakupContacts?.phone ? result.gosZakupContacts.phone.map((item) => item.value).join("; ") : "",
    result.basicInfo.registrationDate.value,
    result.basicInfo.primaryOKED.value,
    result.basicInfo.secondaryOKED.value?.join("; "),
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
    const startColumn = 6; // Assuming the Date column is the 6th column (index 5)
    const startRow = selectedRange.rowIndex;

    // Insert headers
    const headerRange = sheet.getRangeByIndexes(0, startColumn, 1, headers.length);
    headerRange.values = [headers];

    headerRange.copyFrom(sheet.getRange("A1"), "Formats");

    // Insert data
    const dataRange = sheet.getRangeByIndexes(startRow, startColumn, mappedData.length, mappedData[0].length);
    dataRange.values = mappedData;

    dataRange.format.autofitColumns();

    // Set fixed width and enable text wrapping for 'primary' and 'secondary' columns
    const fixedWidthColumns = [1, 4, 5]; // Indexes of 'address', 'primary', and 'secondary' columns
    const fixedWidth = 400;

    fixedWidthColumns.forEach((index) => {
      const columnRange = sheet.getRangeByIndexes(startRow, startColumn + index, mappedData.length + 1, 1);
      columnRange.format.columnWidth = fixedWidth;
      columnRange.format.wrapText = true;
    });

    await context.sync();
  });
}
