!function(){"use strict";var e,t,n,o,a={14385:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},98362:function(e,t,n){e.exports=n.p+"assets/logo-filled.png"},58394:function(e,t,n){e.exports=n.p+"e29b7e77799592cd5fbb.css"}},r={};function s(e){var t=r[e];if(void 0!==t)return t.exports;var n=r[e]={exports:{}};return a[e](n,n.exports,s),n.exports}s.m=a,s.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return s.d(t,{a:t}),t},s.d=function(e,t){for(var n in t)s.o(t,n)&&!s.o(e,n)&&Object.defineProperty(e,n,{enumerable:!0,get:t[n]})},s.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),s.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;s.g.importScripts&&(e=s.g.location+"");var t=s.g.document;if(!e&&t&&(t.currentScript&&(e=t.currentScript.src),!e)){var n=t.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!e||!/^http(s?):/.test(e));)e=n[o--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),s.p=e}(),s.b=document.baseURI||self.location.href,function(){async function e(){const e=document.getElementById("enrichData");e.disabled=!0,e.classList.add("disabled");try{await Excel.run((async e=>{const t=e.workbook.getSelectedRange();t.load("rowIndex, rowCount"),await e.sync();const n=t.rowIndex+1,o=`A${n}:F${n+t.rowCount-1}`,a=e.workbook.worksheets.getActiveWorksheet().getRange(o);a.load("values"),await e.sync();const r=a.values.map((e=>e[2])),s=await async function(e){const t={isDeleted:!1,registrationDate:{value:null},onMarket:null,ceo:{value:{title:""}},primaryOKED:{value:""},secondaryOKED:{value:null},addressRu:{value:""}},n=e.map((async e=>{if(!e)return{basicInfo:t,gosZakupContacts:null};const n=await fetch(`https://apiba.prgapp.kz/CompanyFullInfo?id=${e}&lang=ru`,{headers:{accept:"*/*","accept-language":"en-GB,en-US;q=0.9,en;q=0.8,ru;q=0.7",priority:"u=1, i","sec-ch-ua":'"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',"sec-ch-ua-mobile":"?0","sec-ch-ua-platform":'"Windows"',"sec-fetch-dest":"empty","sec-fetch-mode":"cors","sec-fetch-site":"cross-site",Referer:"https://ba.prg.kz/","Referrer-Policy":"strict-origin-when-cross-origin"},body:null,method:"GET"});if(!n.ok)throw new Error(`Error fetching data for ID ${e}: ${n.statusText}`);return await n.json()}));return(await Promise.allSettled(n)).map((e=>{if("fulfilled"!==e.status||e.value instanceof Error){const n="rejected"===e.status?e.reason:e.value;return{basicInfo:{...t,ceo:{value:{title:n.message??n}}},gosZakupContacts:null}}return e.value}))}(r);await async function(e,t){const n=e.map((e=>{const t=function(e){if(!e)return null;const t=(e=e.replace(/Г\.А\.,\s*/g,"").toLowerCase()).match(/(?:г\.|город)\s*([а-яё\s]+)/),n=t?t[1].trim().toUpperCase():"",o=e.match(/([а-яё\s]+(?:область))/);return{city:n,region:o?o[1].trim().toUpperCase():""}}(e.basicInfo.addressRu.value);return[e.basicInfo.isDeleted?"Организация удалена в источниках":e.basicInfo.ceo.value?.title??"",e.basicInfo.addressRu.value??"",t?.city?t.city:t?.region??"",e.gosZakupContacts?.phone?e.gosZakupContacts.phone.map((e=>e.value)).join("; "):"",e.basicInfo.registrationDate.value??"",e.basicInfo.primaryOKED.value??"",e.basicInfo.secondaryOKED.value?.join("; ")??""]})),o=["name","address","city","phone","registration","primary","secondary"];await Excel.run((async e=>{const a=e.workbook.worksheets.getActiveWorksheet(),r=a.getRange(t);r.load("columnIndex, rowIndex, rowCount"),await e.sync();const s=r.rowIndex,c=a.getRangeByIndexes(0,6,1,o.length);c.values=[o],c.copyFrom(a.getRange("A1"),"Formats");const i=a.getRangeByIndexes(s,6,n.length,n[0].length);i.values=n,i.format.autofitColumns();[1,5,6].forEach((e=>{const t=a.getRangeByIndexes(s,6+e,n.length+1,1);t.format.columnWidth=400,t.format.wrapText=!0})),await e.sync()}))}(s,o)}))}catch(e){console.error(e)}finally{e.disabled=!1,e.classList.remove("disabled")}}async function t(){try{await Excel.run((async e=>{const t=e.workbook.worksheets.getActiveWorksheet();t.getUsedRange().load("values"),await e.sync();const a="по подаче",r="по цене",s="по адресу";await n(a,e),await n(r,e),await n(s,e),await o(t,s,8,e),await o(t,r,3,e),await o(t,a,5,e),await e.sync()}))}catch(e){console.error(e)}}async function n(e,t){const n=t.workbook.worksheets.getItemOrNullObject(e);await t.sync(),n.isNullObject||(n.delete(),await t.sync())}async function o(e,t,n,o){const a=e.copy(Excel.WorksheetPositionType.after,e);a.name=t,a.getUsedRange().getResizedRange(-1,0).getOffsetRange(1,0).sort.apply([{key:n,ascending:!0}]);try{await o.sync()}catch(e){console.log(e)}}Office.onReady((n=>{if(n.host===Office.HostType.Excel){const n=document.getElementById("app-body");n&&(n.style.display="flex");const o=document.getElementById("enrichData");o&&(o.onclick=e);const a=document.getElementById("createViews");a&&(a.onclick=t)}}))}(),e=s(14385),t=s.n(e),n=new URL(s(58394),s.b),o=new URL(s(98362),s.b),t()(n),t()(o)}();
//# sourceMappingURL=taskpane.js.map