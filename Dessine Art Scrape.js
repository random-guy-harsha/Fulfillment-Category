
'Finds the first product and get the link

function extractProductItemURLs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("B2:B247").getValues();

  for (var i = 0; i < data.length; i++) {
    var url = data[i][0];
    var extractedURL = null;

    if (url) {
      var response = UrlFetchApp.fetch(url);
      var content = response.getContentText();

      var regexPattern = /<a[^>]+?href="\/products\/([^"]+)"/;
      var match = content.match(regexPattern);

      if (match && match.length > 1) {
        extractedURL = "/products/" + match[1];
      }
    }

    sheet.getRange(i + 2, 3).setValue(extractedURL);
  }
}




'Scrapes div.Rte from all links

function extractInfoFromDivRte() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  for (var i = 246; i <= 247 ; i++) {
    
  var url = sheet.getRange( i, 4).getValues();
  var response = UrlFetchApp.fetch(url);
  var content = response.getContentText();

  var regexPattern = /<div class="Rte"[^>]*>([\s\S]*?)<\/div>/;
  var matches = content.match(regexPattern);

  var extractedInfo = null;

  if (matches && matches.length > 1) {
    extractedInfo = matches[1];
  }

  sheet.getRange(i , 5).setValue(extractedInfo);
}
}



