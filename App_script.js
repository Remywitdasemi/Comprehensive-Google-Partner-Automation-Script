// Global variables
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const CLIENT_ID = 'YOUR_CLIENT_ID_HERE';
const CLIENT_SECRET = 'YOUR_CLIENT_SECRET_HERE';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Partner Console')
    .addItem('Authorize', 'showAuthorizationDialog')
    .addItem('Process New Orders', 'processNewOrders')
    .addItem('Open Order Form', 'openOrderForm')
    .addItem('Generate Report', 'generateReport')
    .addToUi();
}

function showAuthorizationDialog() {
  var authInfo = getOAuthService().getAuthorizationUrl();
  var template = HtmlService.createTemplate(
      '<a href="<?= authInfo ?>" target="_blank">Authorize</a>. ' +
      'Reopen the sidebar when the authorization is complete.');
  template.authInfo = authInfo;
  var page = template.evaluate();
  SpreadsheetApp.getUi().showModalDialog(page, 'Authorize');
}

function getOAuthService() {
  return OAuth2.createService('GooglePartners')
      .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
      .setTokenUrl('https://accounts.google.com/o/oauth2/token')
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)
      .setCallbackFunction('authCallback')
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope('https://www.googleapis.com/auth/apps.order https://www.googleapis.com/auth/admin.directory.user https://www.googleapis.com/auth/admin.directory.domain')
      .setParam('access_type', 'offline')
      .setParam('approval_prompt', 'force');
}

function authCallback(request) {
  var service = getOAuthService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}

function processNewOrders() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Orders');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] === 'Pending') {
      var clientName = data[i][1];
      var domain = data[i][2];
      var email = data[i][3];
      var licenseCount = data[i][4];
      
      try {
        // Register domain
        var domainResult = registerDomain(domain);
        
        // Setup Google Workspace
        var workspaceResult = setupGoogleWorkspace(clientName, domain, email, licenseCount);
        
        // Mark as processed
        sheet.getRange(i + 1, 6).setValue('Processed');
        
        // Send confirmation email
        sendConfirmationEmail(email, clientName, domain);
        
        Logger.log('Processed order for ' + clientName);
      } catch (error) {
        Logger.log('Error processing order for ' + clientName + ': ' + error.message);
        sheet.getRange(i + 1, 6).setValue('Error: ' + error.message);
      }
    }
  }
}

function registerDomain(domain) {
  // Implement domain registration logic here
  // This will depend on your domain registrar's API
  Logger.log('Registering domain: ' + domain);
  return { success: true, message: 'Domain registered successfully' };
}

function setupGoogleWorkspace(clientName, domain, email, licenseCount) {
  var service = getOAuthService();
  
  // Create customer
  var customerInsertUrl = 'https://www.googleapis.com/apps/reseller/v1/customers';
  var customerPayload = {
    customerDomain: domain,
    customerName: clientName
  };
  var customerResponse = UrlFetchApp.fetch(customerInsertUrl, {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + service.getAccessToken()
    },
    contentType: 'application/json',
    payload: JSON.stringify(customerPayload)
  });
  var customerId = JSON.parse(customerResponse.getContentText()).customerId;
  
  // Create subscription
  var subscriptionInsertUrl = 'https://www.googleapis.com/apps/reseller/v1/customers/' + customerId + '/subscriptions';
  var subscriptionPayload = {
    customerId: customerId,
    skuId: 'Google-Apps-For-Business',
    plan: {
      planName: 'ANNUAL_MONTHLY_PAY'
    },
    seats: {
      numberOfSeats: licenseCount
    }
  };
  var subscriptionResponse = UrlFetchApp.fetch(subscriptionInsertUrl, {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + service.getAccessToken()
    },
    contentType: 'application/json',
    payload: JSON.stringify(subscriptionPayload)
  });
  
  return { success: true, message: 'Google Workspace set up successfully' };
}

function sendConfirmationEmail(email, clientName, domain) {
  var subject = 'Your Google Workspace Account is Ready';
  var body = 'Dear ' + clientName + ',\n\n' +
             'Your domain ' + domain + ' has been successfully registered ' +
             'and your Google Workspace account has been set up.\n\n' +
             'If you have any questions, please don't hesitate to contact us.\n\n' +
             'Best regards,\nOnline is Easy Team';
  
  MailApp.sendEmail(email, subject, body);
}

function openOrderForm() {
  var html = HtmlService.createHtmlOutputFromFile('OrderForm')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'New Order Form');
}

function addNewOrder(orderData) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Orders');
  sheet.appendRow([new Date(), orderData.clientName, orderData.domain, orderData.email, orderData.licenseCount, 'Pending']);
  return 'Order added successfully';
}

function generateReport() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Orders');
  var data = sheet.getDataRange().getValues();
  
  var report = "Order Report\n\n";
  
  for (var i = 1; i < data.length; i++) {
    report += data[i][1] + " - " + data[i][2] + " - " + data[i][5] + "\n";
  }
  
  var docName = "Order Report - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var doc = DocumentApp.create(docName);
  doc.getBody().setText(report);
  
  SpreadsheetApp.getUi().alert('Report generated: ' + doc.getUrl());
}
