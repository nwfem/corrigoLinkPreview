/**
* Entry point for an NWFEM Corrigo Work Order link preview
*
* @param {!Object} event
* @return {!Card}
*/
// Creates a function that passes an event object as a parameter.
function workOrderPreview(event) {

  // If the event object URL matches a specified pattern for support case links.
  if (event.docs.matchedUrl.url) {

    // Uses the event object to parse the URL and identify the work order ID.
    const segments = event.docs.matchedUrl.url.split('/');
    const woId = segments[segments.length - 1];

    // Calls Corrigo API to get WorkOrder data
    var workOrder = getCorrigoWorkOrder(woId);

    // Parses Work Order Location
    var loc = workOrder.Data.ShortLocation;
    var locBreak = loc.lastIndexOf(" - ");

    var priority = workOrder.Data.Priority.Id;
    if (priority == 1)
      priority = "Emergency";
    else if (priority == 2)
      priority = "SLC Emergency";
    else if (priority == 3)
      priority = "SLC BUN 4 Hour";
    else if (priority == 4)
      priority = "SLC Urgent";
    else if (priority == 5)
      priority = "SLC Non-Emergency";
    else if (priority == 6)
      priority = "SLC Planned";
    else if (priority == 7)
      priority = "SLC Project";
    else if (priority == 8)
      priority = "Project";
    else if (priority == 9)
      priority = "Planned";
    else if (priority == 10)
      priority = "Urgent";
    else if (priority == 11)
      priority = "Non-Emergency";

    // Builds a preview card with the case ID, title, and description
    const woHeader = CardService.newCardHeader()
      .setTitle(`${workOrder.Data.Number}`);
    
    const woCust = CardService.newDecoratedText()
      .setTopLabel("Customer")
      .setText(loc.slice(locBreak + 3));
    
    const woStore = CardService.newDecoratedText()
      .setTopLabel("Store")
      .setText(loc.slice(0, locBreak))
      .setWrapText(true);
    
    const woPoNum = CardService.newDecoratedText()
      .setTopLabel("Purchase Order")
      .setText(workOrder.Data.PoNumber);

    const woStatus = CardService.newDecoratedText()
      .setTopLabel("Status")
      .setText(workOrder.Data.StatusId);

    const woPriority = CardService.newDecoratedText()
      .setTopLabel("Priority")
      .setText(priority);

    const woCreation = CardService.newDecoratedText()
      .setTopLabel("Date Created")
      .setText(workOrder.Data.DtCreated.split("T")[0]);

    // Returns the card.
    // Uses the text from the card's header for the title of the smart chip.
    return CardService.newCardBuilder()
      .setHeader(woHeader)
      .addSection(CardService.newCardSection().addWidget(woStore).addWidget(woCust).addWidget(woPoNum))
      .addSection(CardService.newCardSection().addWidget(woCreation).addWidget(woPriority).addWidget(woStatus))
      .build();
  }
}

function getCorrigoWorkOrder(woId) {
  var sp = PropertiesService.getScriptProperties();
  var corrigoName = sp.getProperty('CORRIGO_COMPANY_NAME');

  // Get token/company hostname using Corrigo Authentication API
  var token = getCorrigoToken(sp.getProperty('CORRIGO_CLIENT_ID'), sp.getProperty('CORRIGO_CLIENT_SECRET'));
  var hostname = getCorrigoHostname(corrigoName, token);

  var url = `${hostname}api/v1/base/WorkOrder/${woId}`;
  var options = {
    'headers': {
      "CompanyName": corrigoName,
      "Authorization": "Bearer " + token
    }
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response);
  return JSON.parse(response.getContentText());
}

// Get Corrigo Enterprise Company Token using API
function getCorrigoToken(clientID, clientSecret) {
  var url = "https://oauth-pro-v2.corrigo.com/OAuth/token";
  var options = {
    'method': 'post',
    'payload': {
      'grant_type': 'client_credentials',
      'client_id': clientID,
      'client_secret': clientSecret
    },
    'contentType': 'application/x-www-form-urlencoded'
  };

  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response).access_token;
}

// Get Corrigo Enterprise Company Hostname using API
function getCorrigoHostname(companyName, token) {
  var url = "https://am-apilocator.corrigo.com/api/v1/cmd/GetCompanyWsdkUrlCommand";
  var options = {
    'method': 'post',
    'headers': {
      "CompanyName": companyName,
      "Authorization": "Bearer " + token
    },
    'contentType': "application/json",
    'payload': JSON.stringify({
      "Command": {
        "ApiType": "REST",
        "CompanyName": companyName,
        "Protocol": "HTTPS"
      }
    })
  };

  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText()).CommandResult.Url;
}