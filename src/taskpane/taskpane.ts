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
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");
      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      await fetchData(["580527300014", "550619300445"]);
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

async function fetchData(ids: string[]) {
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

    populateExcel(results);
  } catch (error) {
    console.error("Error fetching data:", error);
  }
}

async function populateExcel(data: CompanyFullInfo[]) {
  // Map the data to the desired columns
  const mappedData = data.map((result) => [
    result.basicInfo.ceo.value.title,
    result.basicInfo.addressRu.value,
    result.gosZakupContacts?.phone.length > 0 ? result.gosZakupContacts.phone.map((item) => item.value).join("; ") : "",
    result.basicInfo.registrationDate.value,
    result.basicInfo.primaryOKED.value,
    result.basicInfo.secondaryOKED.value.join("; "),
  ]);

  // Add headers
  const headers = ["name", "address", "phone", "registration", "primary", "secondary"];
  mappedData.unshift(headers);

  // Populate the Excel sheet
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(`A1:F${mappedData.length}`);
    range.values = mappedData;
    await context.sync();
  });
}
