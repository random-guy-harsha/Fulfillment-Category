'HomeDecorzStore image link scraping

function getImageLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Input");

  var lastRow = sourceSheet.getLastRow();
  var output = [];

  for (var row = 2; row <= lastRow; row++) {
    var url = sourceSheet.getRange(row, 1).getValue();

    var html = UrlFetchApp.fetch(url).getContentText();
    var regex = /https:\/\/sp-ao\.shortpixel\.ai\/client\/to_auto,q_lossy,ret_img,w_1200,h_800\/(https:\/\/homedecorzstore\.com\/wp-content\/uploads\/\d{4}\/\d{2}\/[^"']+\.(?:jpg|jpeg|png|gif))/gi;

    var matches = html.match(regex);
    var uniqueLinks = matches ? [...new Set(matches)].join(", ") : "No image links found.";

    output.push([url, uniqueLinks]);
  }

  var targetSheet = ss.insertSheet("Output");
  var range = targetSheet.getRange(1, 1, output.length, output[0].length);
  range.setValues(output);
}
