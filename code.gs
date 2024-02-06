var currentSheet;
function getActiveSheet()
{
  if (currentSheet != null) return currentSheet
  else currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return currentSheet;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Custom Menu');
  menu.addItem('Fetch Campaigns', 'fetchMailchimpCampaigns').addToUi();

  currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function getMcApiKey()
{
  return 'ENTER YOUR API KEY';
}

function getListId()
{
  return 'ENTER YOUR LIST ID';
}

function getCellToTheRight() {
  var currentCell = currentSheet.getActiveCell();
  
  // Offset by 0 rows and 1 column to the right
  var cellToTheRight = currentCell.offset(0, 1);
  
  return cellToTheRight;  // This is a Range object representing the cell to the right
}

/* fetch campaigns */

function fetchMailchimpCampaigns() {
  var currentSheet = getActiveSheet();
  var sheet = currentSheet;

//use this later if possible
  var startDate = currentSheet.getRange('B1').getDisplayValue();
  var endDate = currentSheet.getRange('B2').getDisplayValue();
  var count = 1000; //max allowed by mailchimp api

  var apiEndpoint = 'https://us9.api.mailchimp.com/3.0/campaigns';
  var apiKey = getMcApiKey();
  var params = {
    'method': 'get',
    'headers': {
      'Authorization': 'apikey ' + apiKey
    },
    'muteHttpExceptions': true
  };

  // Adding query parameters for filtering by date range
  apiEndpoint += '?since_send_time=' + startDate + '&before_send_time=' + endDate + '&count=' + count + '&sort_field=send_time';
  
  var response = UrlFetchApp.fetch(apiEndpoint, params);
  var json = JSON.parse(response.getContentText());    

  // Log the subject lines or other details
  json.campaigns.forEach(function(campaign) {
    Logger.log('Subject: ' + campaign.settings.subject_line);
    Logger.log(campaign.settings);

    var currentCell = sheet.getActiveCell();
    var col = currentCell.getColumn();
    var row = currentCell.getRow();

    fillCampaign(campaign, currentCell);
    
    //logic to move to next row, starting column
    sheet.setActiveSelection(sheet.getRange(row+1,col));
  });
}

function fillCampaign (campaign, currentCell)
{
  // Assuming you want to write details to the current cell (or you can modify the target as needed)
  var targetCell = currentCell;
  currentCell.setValue(campaign['id'])
  getCellToTheRight().setValue(campaign.settings['subject_line']).activate();
  getCellToTheRight().setValue(campaign.settings['title']).activate();
  getCellToTheRight().setValue(campaign.settings['from_name']).activate();
  getCellToTheRight().setValue(campaign['archive_url']).activate();
  getCellToTheRight().setValue(campaign['send_time']).activate();
  getCellToTheRight().setValue(campaign['emails_sent']).activate();
  getCellToTheRight().setValue(campaign.report_summary['opens']).activate();
  getCellToTheRight().setValue(campaign.report_summary['unique_opens']).activate();
  getCellToTheRight().setValue(campaign.report_summary['open_rate']).activate();
  getCellToTheRight().setValue(campaign.report_summary['clicks']).activate();
  getCellToTheRight().setValue(campaign.report_summary['subscriber_clicks']).activate();
  getCellToTheRight().setValue(campaign.report_summary['click_rate']).activate();
}
