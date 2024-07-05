import { enrichData } from "./enrich-data";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const appBody = document.getElementById("app-body");
    if (appBody) {
      appBody.style.display = "flex";
    }

    const enrichButton = document.getElementById("enrichData");
    if (enrichButton) {
      enrichButton.onclick = enrichData;
    }
  }
});
