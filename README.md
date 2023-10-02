# function doGet() {
  let filter = `{subject:"You've Made A Sale"}`;
  let spreadsheet = SpreadsheetApp.openByUrl(getSheetUrl());
  let values = spreadsheet.getSheetByName('Order').getRange(2, 1, 1).getValues();
  if (values[0][0] != '') {
    filter = filter + ' AND after:'+ getYesterday();
  }
  let sheetUrl = getSheetUrl();
  let sheetName = 'Order';
  let messages = getRelevantMessages(filter);
  parseMessageData(messages, sheetUrl, sheetName);
}
function getSheetUrl() {
  let SS = SpreadsheetApp.getActiveSpreadsheet();
  let ss = SS.getActiveSheet();
  let url = '';
  url += SS.getUrl();
  url += '#gid=';
  url += ss.getSheetId();
  return url;
}
function getYesterday() {
  let yesterday = new Date(Date.now() - 86400000);
  let dd = String(yesterday.getDate()).padStart(2, '0');
  let mm = String(yesterday.getMonth() + 1).padStart(2, '0');
  let yyyy = yesterday.getFullYear();
  return yyyy + '/' + mm + '/' + dd;
}
function getSheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}
function getRelevantMessages(filter) {
  let threads = GmailApp.search(filter);
  let messages = [];
  threads.forEach(function(thread) {
      messages.push(thread.getMessages()[0]);
  });
  return messages.reverse();
}
function prependRow(sheet, rowData) {
  sheet.insertRowBefore(2).getRange(2, 1, 1, rowData.length).setValues([rowData]);
}
function saveDataToSheetCurrent(record, sheetUrl, sheetName) {
  let spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  let sheet = spreadsheet.getSheetByName(sheetName);
  // sheet.appendRow(record);
  prependRow(sheet, record);

}
function getTimesFromSheet(sheetUrl, sheetName, titles) {
  let spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  let sheet = spreadsheet.getSheetByName(sheetName);
  let times = [];
  if (!sheet) {
    let newSheet = spreadsheet.insertSheet();
    newSheet.setName(sheetName);
    newSheet.appendRow(titles);
  } else {
    let values = sheet.getDataRange().getValues();
    values.forEach(function(row, key) {
      if (key == 0 || row[0] == '') {
          return;
      }
      times.push(row);
    });
  }
  return times;
}
function extractData(data, startStr, endStr) {
  let startIndex, endIndex, text = '';
  startIndex = data.indexOf(startStr);
  if (startIndex != -1) {
      startIndex += startStr.length;
      text = data.substring(startIndex);
      if (endStr) {
          endIndex = text.indexOf(endStr);
          if (endIndex != -1) {
              text = text.substring(0, endIndex);
          } else {
              text = '';
          }
      }
  }
  return text;
}
function getItemsToCeil(text, before, after) {
  let string = '';
  let textOveride = text;
  let isCheck = true;
  while (isCheck) {
      let extract = extractData(textOveride, before, after).trim();
      if (string == '') {
          string = extract;
      } else {
          string = string + "\n" + extract;
      }
      textOveride = textOveride.replace(before, "");
      isCheck = textOveride.includes(before);
  }
  return string;
}
function getOrderFromSheet(sheetUrl, sheetName) {
  let spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  let sheet = spreadsheet.getSheetByName(sheetName);
  let values = sheet.getDataRange().getValues();
  let orders = [];
  values.forEach(function(row, key) {
    if (key == 0 || row[0] == '') {
        return;
    }
    if (key == 200) {
      return;
    }
    orders.push(row[4]);
  });
  if (values[0][0] == '') {
      sheet.appendRow([
          "Stt",
          "Time",
          "Date",
          "Account",
          "Order",
          "Etsy Account",
          "Owner",
          "Size",
          "Country",
          "Margin",
          "Total",
          "USDconvert($)",
          "Product",
          "Nameproduct",
          "Estimateprofit",
      ]);
  }
  return orders;
}
function checkLiveEtsy(shopLink) {
  try {
    if(!shopLink) return '';
    const content = UrlFetchApp.fetch(shopLink).getContentText();
    let status = ''
    if (content.includes(' found a glitch.')) {
      status = 'SUSPENDED';
    }  else {
      status = 'LIVE';
    }
    return status;
  } catch (err) {
    return 'SUSPENDED';
  }
}
function checkLive(sheetUrl, title) {
  createSheet(sheetUrl, "Account", title);
  let spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  let sheet = spreadsheet.getSheetByName("Account");
  sheet.getRange('C2').setValue('=checkLiveEtsy(B2)');
}
function parseMessageData(messages, sheetUrl, sheetName) {
  try {
     let orders = getOrderFromSheet(sheetUrl, "Order");
    console.log(messages.length);
    for (let m = 0; m < messages.length; m++) {
        const text = messages[m].getPlainBody();
        const bodyHTML = messages[m].getBody();
        const subject = messages[m].getSubject();
        let date = messages[m].getDate();
        let orders = getOrderFromSheet(sheetUrl, "Order");
        let time = Utilities.formatDate(date, "Asia/Ho_Chi_Minh", "HH:mm : ss a").toString();
        date = Utilities.formatDate(date, "Asia/Ho_Chi_Minh", "dd/MM/yyyy").toString();
        let etsyAccount = messages[m].getTo();
        let nameShop = getItemsToCeil(bodyHTML, "Shop:", "\n");
        if (nameShop == '') {
          nameShop = getItemsToCeil(bodyHTML, "shop:", "\n");
        }
        let order = extractData(subject, "You've Made A Sale -", " (").trim(); 
        let margin = extractData(subject, "(", ")").trim();
        let price = getItemsToCeil(bodyHTML, "Total price</strong>:", "</p>");
        let country = getItemsToCeil(bodyHTML, "an admirer of art in ", " picked your");
        let size = getItemsToCeil(text, "Size:", "\n");
        let owner = getItemsToCeil(bodyHTML, "<strong>Hi", ",</strong>").trim();
        let margin2 = extractData(subject, "(US$", ")").trim();
            let rate=1;
            let tax =  7.1;
            if (margin2 == '') {
              margin2 = extractData(subject, "(£", ")").trim();
              rate = 1.25; tax =  7.1;
            }
            if (margin2 == '') {
              margin2 = extractData(subject, "(€", ")").trim();
              rate = 1.07; tax =  19.5;
            }  
            if (margin2 == '') {
              margin2 = extractData(subject, "(CA$", ")").trim();
              rate = 0.74; tax =  12.5;
            }  
            if (margin2 == '') {
              margin2 = extractData(subject, "(AU$", ")").trim();
              rate = 0.66; tax =  1;
            }  
        let convertUSD = margin2 * rate;
        let product = getItemsToCeil(bodyHTML, "1x ", "of");
        if (product == '') {
          product = getItemsToCeil(bodyHTML, "2x ", "of");
        }
         if (product == '') {
          product = getItemsToCeil(bodyHTML, "3x ", "of");
        }
         if (product == '') {
          product = getItemsToCeil(bodyHTML, "4x ", "of");
        }
         if (product == '') {
          product = getItemsToCeil(bodyHTML, "5x ", "of");
        }
         if (product == '') {
          product = getItemsToCeil(bodyHTML, "6x ", "of");
        }
         if (product == '') {
          product = getItemsToCeil(bodyHTML, "7x ", "of");
        }
         if (product == '') {
          product = getItemsToCeil(bodyHTML, "8x ", "of");
        }
         if (product == '') {
          product = getItemsToCeil(bodyHTML, "9x ", "of");
        }
         if (product == '') {
          product = getItemsToCeil(bodyHTML, "10x ", "of");
        }
         if (product == '') {
          product = getItemsToCeil(bodyHTML, "11x ", "of");
        }
        let nameproduct = getItemsToCeil(bodyHTML, "<strong>1x", "</strong>").replace("of", ":1x");
         if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>2x", "</strong>").replace("of", ":2x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>3x", "</strong>").replace("of", ":3x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>4x", "</strong>").replace("of", ":4x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>5x", "</strong>").replace("of", ":5x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>6x", "</strong>").replace("of", ":6x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>7x", "</strong>").replace("of", ":7x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>8x", "</strong>").replace("of", ":8x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>9x", "</strong>").replace("of", ":9x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>10x", "</strong>").replace("of", ":10x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>11x", "</strong>").replace("of", ":11x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>12x", "</strong>").replace("of", ":12x");
        }
        if (nameproduct == '') {
          nameproduct = getItemsToCeil(bodyHTML,"<strong>13x", "</strong>").replace("of", ":13x");
        }
        let estimateprofit = convertUSD - (convertUSD * tax / 100);
        let isCheck = false;
    orders.forEach(function(value) {
      if (value == order) {
          isCheck = true;
          return;
      }
    });
        if (price == '' && owner == '' || isCheck ) {
          continue;
        }
        saveDataToSheetCurrent([
          '=ROW() - 1',
            time,
          date,
          etsyAccount,
          order,
          owner,
          product,
          nameproduct,
          size,
          country,
          margin,
          price,
          convertUSD,
          estimateprofit, 
        ], sheetUrl, sheetName);
    }
  } catch (err) {
    console.log(err.message);
  }
}
