/**
 * OpenBID Toolkit Add-on
 * This Google Sheets add-on is designed to work exclusively with "OpenBID" spreadsheet templates.
 * It will create copies of template files. Download and insert images, insert YouTube chapters, enable OpenAI functions. 
 * It will share an NDA, generate and share a copy of the OpenBID Vendor template for the corresponding episode to the selected vendor.
 * It will prep the sheet for sharing and create backup files for PDF export.
 * This Google Sheets add-on also checks if the user's email is in a list of authorised users.
 * If the user is not in the list, a popup with a link to sign up is displayed instead.
 */


// The latest template files
const OPENBIDTEMPLATECLIENT = "1YhYsxJINHLTZ6ymRrtoAeJ5EyuNZt0wm8gp9JbAZYNg"; // OpenBID Client
const OPENBIDTEMPLATEVENDOR = "1K28U0_Qa5F_wlcsIS0gLH8Qa_vojlBFG8IIZvyHXGnY"; // OpenBID Vendor
const OPENBIDTEMPLATEPROFITPLANNER = "1w8rinFc-5wu20dCwc3Tgydse7fiskwqX205c3fzewZI"; //OpenBID Vendor Profit Planner
const OPENBIDTEMPLATEOSCR = "1Lk3TbDvQ0-IInh1Cd5LHWzuPRMnDumoHlfooD8-wjmc"; //OpenBID OSCR

// Set to 'true' to enable subscription check, 'false' to disable
const ENABLE_SUBSCRIPTION_CHECK = false;


// The columns in the Overview sheet with the data needed
const PROJECT_NAME_COLUMN = 3; // The project name for that episode / block
const EPISODE_COLUMN = 2; // The episode / block
const PACKAGE_COLUMN = 9; // The bid package number
const BID_MATERIAL = 14; // The link with the bid material to share


// The location of the NDA data
const ROWNDAOFFSET = 7;  // The first row containing the vendor email addresses
const COLUMNNDAOFFSET = 7;  // The column containing the vendor email addresses


// The location of the Bid data
const ROWBIDOFFSET = 4;  // The first row containing the episode names
const COLUMNBIDOFFSET = 17;  // The first column containing the vendor email addresses


// The offset between the vendor emails on the OVERVIEW tab and the SHOT BID BREAKDOWN and ASSET BID BREAKDOWN
const VENDOR_EMAIL_OFFSET = 13;


// The column in the bid summary with the bid urls on the OVERVIEW tab
const ALLBIDURLS = 2;


// The folder name to use
const OPENBID_ROOT_FOLDER = "OpenBID Toolkit Files";


// The column to put the YouTube chapter link into
const PUTYOUTUBELINK = 5;


// The ID of the OpenBID Vendor Google Sheet
const OPENBIDVENDOR_SHEET_ID = "1rxKIHWdiRFcvnKlGDqoievy7MGtM1SgzqJHLKixT5nY";


// The ID of the Google Sheet containing the 'PaidUsers' named range
const SUBSCRIBER_SHEET_ID = "1aczl4sLwAnc_zFb4IYWwNQk0ygjX8yEDXallQhkfyf0";


// The ID of the Google Sheet containing the user logs
const USERLOG_SHEET_ID = "1_LIzxw61jn4S6DHcFM8E2fjvSfl6seHqG4QM8GeYQnQ";


/**
 * Adds a custom menu to the active spreadsheet when the add-on is first installed.
 * Calls the onOpen function to add the custom menu after installation.
 */
function onInstall() {
  onOpen(); // Call onOpen to add the custom menu after installation
}


/**
 * Adds a custom menu to the active spreadsheet when it is opened.
 * Calls the addCustomMenu function to add the custom menu.
 */
function onOpen() {
  addCustomMenu();
}


/**
 * Adds a custom menu to the active spreadsheet with the specified menu items.
 * The custom menu includes the options specified.
 */
function addCustomMenu() {
  var ui = SpreadsheetApp.getUi();

  // Create the custom menu
  var menu = ui.createMenu('OpenBID Toolkit')
      .addItem('VFX Vendor Availability for Your Show', 'openGetVendorFormInPopup')
      .addSeparator()
      .addItem('Create an OpenBID Client Sheet', 'createClientOpenbid')
      .addItem('Create an OpenBID Vendor Sheet', 'createVendorOpenbid')
      .addItem('Create an OpenBID Vendor Portfolio Sheet', 'createProfitPlannerOpenbid')
      .addItem('Create an OpenBID On-Set Camera Report Sheet', 'createCameraReportOpenbid')
      .addSeparator()
      .addItem('Send NDAs to Vendors', 'showNDAPopup')
      .addSeparator()
      .addItem('Install GPT() and DALLE() Functions', 'showKeyPopup')
      .addItem('Download DALL-E Images to Google Drive', 'downloadImages')
      .addItem('Insert Thumbnail Images from Google Drive', 'showImageSidebar')
      .addItem('Insert YouTube Chapters for Shots', 'showChapterSidebar')
      .addSeparator()
      .addItem('Send Bid Breakdowns to Vendors', 'showBidPopup')
      .addSeparator()
      .addItem('Convert Current Shot to Template', 'createTemplateShot')
      .addItem('Version Up Your Bid', 'createNewBid')
      .addSeparator()
      .addItem('Prep Sheet to Share with Client', 'prepBid')
      .addItem('Undo Sheet Prep for Client', 'undoPrepBid')
      .addItem('Create PDF Bid for Client', 'bidExportPdf')
      .addSeparator()
      .addItem('Export to ShotGrid', 'downloadSheetsAsCSV')
      .addSeparator()
      .addItem('Become a Vetted VFX Vendor', 'openVendorApplicationFormInPopup')
      .addSeparator()
      .addItem('Vote for the Next OpenBID Feature', 'openFormInPopup');

  // Add the custom menu to the user interface
  menu.addToUi();
}


/**
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 *
 * Make OpenBID files from templates functions
 * 
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 */


/**
 * createClientOpenbid
 * 
 * This function acts as a wrapper for the 'createOpenbid' function, 
 * specifically for creating a new OpenBID document of the 'client' type. 
 * It calls 'createOpenbid' with the argument 'client'.
 */
function createClientOpenbid() {
  createOpenbid('client');
}

/**
 * createVendorOpenbid
 * 
 * This function acts as a wrapper for the 'createOpenbid' function, 
 * specifically for creating a new OpenBID document of the 'vendor' type.
 * It calls 'createOpenbid' with the argument 'vendor'.
 */
function createVendorOpenbid() {
  createOpenbid('vendor');
}

/**
 * createProfitPlannerOpenbid
 * 
 * This function acts as a wrapper for the 'createOpenbid' function, 
 * specifically for creating a new OpenBID document of the 'profitPlanner' type.
 * It calls 'createOpenbid' with the argument 'profitPlanner'.
 */
function createProfitPlannerOpenbid() {
  createOpenbid('profitPlanner');
}

/**
 * createCameraReportOpenbid
 * 
 * This function acts as a wrapper for the 'createOpenbid' function, 
 * specifically for creating a new OpenBID document of the 'cameraReport' type.
 * It calls 'createOpenbid' with the argument 'cameraReport'.
 */
function createCameraReportOpenbid() {
  createOpenbid('cameraReport');
}


/**
 * OpenBID Spreadsheet Creation Function
 * 
 * This function is designed to streamline the creation of new OpenBID spreadsheets
 * within Google Sheets by utilizing predefined templates. Each template type—Client,
 * Vendor, and Profit Planner—serves a specific purpose in the OpenBID workflow.
 * 
 * Based on user selection, this function creates a copy of the respective template,
 * places it in the appropriate Google Drive folder, and provides the user with a
 * link to the newly created spreadsheet. This automation facilitates efficient
 * management and organization of OpenBID documents, reducing manual effort and
 * minimizing errors.
 *
 * @param {string} openbidType - Specifies the type of OpenBID file to be created.
 */
function createOpenbid(openbidType) {
  // Log user activity
  logUserActivity("Creating a OpenBID file");

  // Access the currently active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Display a temporary notification to inform the user of the file creation
  ss.toast("Creating your file now.", "🥁 Drumroll...", 45);

  // Create or locate the folders where the new OpenBID files will be stored
  var projectFolder = findOrCreateFolder("", "OpenBID Sheets from Latest Templates");

  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" or "PaidUsers" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !(userSubscriptions.includes("PaidSups") || userSubscriptions.includes("PaidUsers"))) {
    handleAccessDenied();
  } else {
    // User is a "PaidSups" or "PaidUsers" subscriber or the subscription checking is not enabled.

    // Determine the template ID based on the specified OpenBID file type
    var openbidTemplateID;
    if (openbidType == "client") {
      openbidTemplateID = OPENBIDTEMPLATECLIENT;
      
      // Log user activity
      logUserActivity("Created OpenBID Client");
    }
    else if (openbidType == "vendor") {
      openbidTemplateID = OPENBIDTEMPLATEVENDOR;

      // Log user activity
      logUserActivity("Created OpenBID Vendor");
    }
    else if (openbidType == "profitPlanner") {
      openbidTemplateID = OPENBIDTEMPLATEPROFITPLANNER;

      // Log user activity
      logUserActivity("Created OpenBID Portfolio");
    }
    else if (openbidType == "cameraReport") {
      openbidTemplateID = OPENBIDTEMPLATEOSCR;

      // Log user activity
      logUserActivity("Created OpenBID OSCR");
    }
    else {
      // Log an error message if an invalid OpenBID file type is provided
      Logger.log("Invalid openbidType: " + openbidType);
      return;
    }

    // Retrieve the specified template file from Google Drive
    var templateFile = DriveApp.getFileById(openbidTemplateID);
    
    // Create a copy of the template file
    var newOpenbidFile = templateFile.makeCopy();
    
    // Retrieve the ID of the new OpenBID file
    var newOpenbidFileID = DriveApp.getFileById(newOpenbidFile.getId());
    
    // Move the new OpenBID file to the designated project folder
    Drive.Files.update(
      { 'parents': [ { 'id': projectFolder.getId() } ] }, 
      newOpenbidFileID.getId()
    );

    // Retrieve the URL of the new OpenBID file
    var newOpenbidFileUrl = newOpenbidFile.getUrl();

    // Display a popup window to provide the user with a link to the new OpenBID file
    showNewFilePopup(newOpenbidFileUrl);
  }
}


/**
 * Displays a popup to inform the user about the created OpenBID spreadsheet.
 * The popup contains a red button that links to the new spreadsheet URL.
 *
 * @param {string} newOpenbidFileUrl - URL of the created OpenBID spreadsheet.
 */
function showNewFilePopup(newOpenbidFileUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Create a popup with a red button linked to the new SpreadsheetUrl using the HTML file
  var htmlTemplate = HtmlService.createTemplateFromFile('newFilePopup');
  htmlTemplate.newOpenbidFileUrl = newOpenbidFileUrl;
  var htmlOutput = htmlTemplate.evaluate().setWidth(450).setHeight(250);

  // Show the popup
  ui.showModalDialog(htmlOutput, 'Created your OpenBID file!');
}


/**
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 *
 * OpenAI functions
 * 
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 */


/**
 * This function creates and displays a modal dialog using an HTML file
 * named 'KeyPopup'. The modal dialog is displayed within the Google Sheets
 * user interface, allowing the ability to enter their OpenAI secret key.
 */
function showKeyPopup() {
  // Log user activity as "Open AI key added"
  logUserActivity("Open AI key pop up opened");

  // Create an HTML output from the 'KeyPopup' HTML file
  var htmlOutput = HtmlService.createHtmlOutputFromFile('KeyPopup')
    .setWidth(540)
    .setHeight(650);

  // Display the modal dialog within the Google Sheets user interface
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'OpenBID Toolkit - Initialise OpenAI');
}


/**
 * This function stores the OpenAI Secret Key in the Google Sheets script Document Properties.
 * @param {string} key - The OpenAI Secret Key to be stored.
 */
function storeKey(key) {
  PropertiesService.getDocumentProperties().setProperty('OPENAI_SECRET_KEY', key);

  // Call function to get the joke
  tellJoke();

  // Log user activity as "Open AI key added"
  logUserActivity("Open AI key added");
}


/**
 * This function gets the OpenAI secret key from Document Properties, then uses it
 * to call the GPT function and show a toast notification in the active spreadsheet.
 */
function tellJoke() {
  // Set up the parameters for the GPT function
  var temperature = 2;
  var prompt = "Tell me a short joke";

  // Call the GPT function to get the joke
  var joke = GPT(temperature, prompt, false);

  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Show the joke in a toast notification
  ss.toast(joke,"🤖",10);
}


/**
 * This function generates text using the GPT-4o model from OpenAI's API.
 *
 * @param {number} modelTemp - The temperature for the GPT model's response. Ranges from 0 to 2, with 2 being the most creative.
 * @param {string} systemPrompt - The system's prompt for the model.
 * @param {string} yourPrompt - The user's prompt for the model.
 * @param {boolean} useCache - A flag that determines whether to use cached responses or not.
 * @return {string} - The response from the model or an empty string if no response was found.
 * 
 * @customfunction
 */
function GPT(modelTemp, systemPrompt, yourPrompt, useCache) {
  // Get the secret key from Document Properties
  var secretKey = PropertiesService.getDocumentProperties().getProperty('OPENAI_SECRET_KEY');

  // Define cache key
  var cacheKey = hash(secretKey + "-" + modelTemp + "-" + yourPrompt);

  if (useCache) {
    // Access the script cache
    var cache = CacheService.getScriptCache();

    // Check if the response is in the cache
    var cachedResponse = cache.get(cacheKey);

    if (cachedResponse != null) {
      // Return the cached response
      return cachedResponse;
    }
  }

  // Define the URL for the OpenAI API
  const url = "https://api.openai.com/v1/chat/completions";

  // Define the payload to be sent in the API request
  const payload = {
    model: "gpt-4o",
    messages: [{"role": "system", "content": systemPrompt}, {"role": "user", "content": yourPrompt}],
    temperature: modelTemp
  };

  // Define the options for the API request
  const options = {
    contentType: "application/json",
    headers: { Authorization: "Bearer " + secretKey },
    payload: JSON.stringify(payload),
  };

  // Send the API request and parse the response as JSON
  const responseJson = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());

  // Find the assistant message in the response
  for (var i = 0; i < responseJson.choices.length; i++) {
    var choice = responseJson.choices[i];
    if (choice.message.role === 'assistant') {
      if (useCache) {
        // Add the assistant's response to the cache, with a 6 hr expiration time
        cache.put(cacheKey, choice.message.content.trim(), 21600);
      }
      // Return the assistant's response
      return choice.message.content.trim();
    }
  }

  // If no assistant message is found, return an empty string
  return '';
}


/**
 * Generates images using DALLE model.
 *
 * @param {number} size - Size of the generated image.
 * @param {string} yourPrompt - Prompt to generate an image from.
 * @param {boolean} useCache - Whether to use the cache or not.
 * @return Response returned by DALLE.
 * @customfunction
 */
function DALLE(size, yourPrompt, useCache) {
  // Get the secret key from Document Properties
  var secretKey = PropertiesService.getDocumentProperties().getProperty('OPENAI_SECRET_KEY');

  // Define the cache key based on the parameters
  var cacheKey = secretKey + "-" + size + "-" + yourPrompt;

  // Check if the useCache parameter is true
  if (useCache) {
    // Access the script cache
    var cache = CacheService.getScriptCache();

    // Check if the image URL is in the cache
    var cachedImageUrl = cache.get(cacheKey);

    // If cached image URL is found, return it
    if (cachedImageUrl != null) {
      return cachedImageUrl;
    }
  }

  // Define the URL for the DALLE API
  const url = 'https://api.openai.com/v1/images/generations';

  // Define the payload to be sent in the API request
  const payload = {
    prompt: yourPrompt,
    n: 1,
    size: size,
  };


  // Define the options for the API request
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + secretKey },
    payload: JSON.stringify(payload),
  };

  // Send the API request and parse the response as JSON
  const responseJson = UrlFetchApp.fetch(url, options).getContentText();
  const imageUrl = JSON.parse(responseJson).data[0].url;

  // Check again if the useCache parameter is true
  if (useCache) {
    // Add the image URL to the cache, with a 6 hour expiration time
    cache.put(cacheKey, imageUrl, 21600);
  }

  // Return the generated image URL
  return imageUrl;
}


/**
 * This function downloads images from URLs provided in column AR of the active sheet in Google Spreadsheet.
 * It names each file according to corresponding entries in column AS.
 * Descriptions for each file are retrieved from column AP.
 * These images are saved in a folder named 'Thumbnails' within the project folder (retrieved from the 'Project' named range) in Google Drive.
 * If no images were found, a toast message is displayed. Upon completion, a toast message containing the Drive folder location is displayed.
 */
function downloadImages() {
  // Log user activity as "Images downloaded"
  logUserActivity("Images downloaded");

  // Get the active Spreadsheet and Sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Get the last row with data in column AR (44th column)
  var lastRow = getLastRowWithData(sheet, 44);

  // If no data is found in column AR, show a toast message and exit function
  if (lastRow === 0) {
    ss.toast('Pick a ChatGPT & DALL-E tab to download from.', '⚠️️ No Images Found', 10);
    return;
  }

  // Retrieve image URLs from column AR, corresponding file names from column AS, and descriptions from column AP
  var imageURLs = sheet.getRange(5, 44, lastRow).getValues();
  var filenames = sheet.getRange(5, 45, lastRow).getValues();
  var descriptions = sheet.getRange(5, 42, lastRow).getValues();

  // Get the project name from the named range "Project"
  var projectNameRange = ss.getRangeByName("Project");
  var projectName = projectNameRange.getValue();

  // Get the OpenBID folder
  var openBIDFolder = getOpenBIDFolder();

  // Create the folders  
  var projectFolder = findOrCreateFolder(openBIDFolder, projectName);
  var thumbnailsFolder = findOrCreateFolder(projectFolder, 'Thumbnails');

  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" or "PaidUsers" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !(userSubscriptions.includes("PaidSups") || userSubscriptions.includes("PaidUsers"))) {
    handleAccessDenied();
  } else {
    // User is a "PaidSups" or "PaidUsers" subscriber or the subscription checking is not enabled.

    // Loop through all imageURLs
    for (var i = 0; i < lastRow; i++) {
      if (imageURLs[i][0]) {
        try {
          // If URL is valid, fetch the image and create a blob
          var response = UrlFetchApp.fetch(imageURLs[i][0]);
          var blob = response.getBlob();

          // Name the blob according to the corresponding entry in column AS and add .png extension
          blob.setName(filenames[i][0] + ".png");
          ss.toast("🔨 " + filenames[i][0] + ".png created!");

          // Save the image in the 'Thumbnails' folder
          var file = thumbnailsFolder.createFile(blob);

          // Set the description of the file from column AP
          file.setDescription(descriptions[i][0]);
        } catch (error) {
          // Logger.log('Error fetching URL at row ' + (i+1) + ': ' + imageURLs[i][0] + ' Error: ' + error);
          ss.toast('⚠️️ Error fetching URL at row ' + (i+1))
          // Continue with the next iteration if this URL is bad.
          continue;
        }
      }
    }
  }
  // Show a toast message containing the Drive folder URL after all images are downloaded
  ss.toast("See " + ss.getName() + " " + OPENBID_ROOT_FOLDER + " -> " + projectName + " -> Thumbnails", '✅ All Images Saved', 60);
}


/**
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 *
 * Send out NDA functions
 * 
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 */


/**
 * This function creates and displays a modal dialog using an HTML file
 * named 'NDAPopup'. The modal dialog is displayed within the Google Sheets
 * user interface, providing additional functionality for the OpenBID Toolkit.
 */
function showNDAPopup() {
  // Log user activity as "NDA pop up shown"
  logUserActivity("NDA pop up shown");

  // Create an HTML output from the 'NDAPopup' HTML file
  var htmlOutput = HtmlService.createHtmlOutputFromFile('NDAPopup')
    .setWidth(680)
    .setHeight(550);

  // Display the modal dialog within the Google Sheets user interface
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'OpenBID Toolkit - Send NDAs');
}


/**
 * This function retrieves the vendor list and their corresponding NDA status
 * from the named ranges "vendorList" and "vendorNDAList" in the active
 * spreadsheet. It returns an object containing the vendor list and their NDA
 * status list.
 *
 * @returns {Object} An object containing the vendor list and their NDA status list.
 */
function getVendorNDAData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get the named ranges for vendor list and NDA status list
    var vendorListRange = ss.getRangeByName("vendorList");
    var vendorNDAListRange = ss.getRangeByName("vendorNDAList");

    // Retrieve the values from the named ranges
    var vendorList = vendorListRange.getValues();
    var vendorNDAList = vendorNDAListRange.getValues();

    return {
      vendorList: vendorList,
      vendorNDAList: vendorNDAList
    };
  } catch (error) {
    // Return the error message
    ss.toast("You can only do this from an OpenBID spreadsheet.", "⚠️ Error");
    return {
      error: error.message
    };
  }
}


/**
 * This function sends a Non-Disclosure Agreement (NDA) to a vendor
 *
 * @param {number} row - The row number of the vendor in the "ADD VENDORS" sheet.
 */
function sendNDA(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ADD VENDORS");

  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !userSubscriptions.includes("PaidSups")) {
    handleAccessDenied();
  } else {
    // User is a "PaidSups" subscriber or the subscription checking is not enabled.

    sendNDAToVendor(ss, sheet, row);
    // Logger.log("DONE");
  }
}


/**
 * This function sends a Non-Disclosure Agreement (NDA) to a vendor
 * and logs the activity for subscribed users.
 *
 * @param {Object} ss - The active Google Spreadsheet object.
 * @param {Object} sheet - The active sheet object from the Google Spreadsheet.
 * @param {number} row - The row number of the vendor in the "ADD VENDORS" sheet.
 */
function sendNDAToVendor(ss, sheet, row) {
  // Get values for projectName, firstName, and lastName from named ranges
  var projectName = ss.getRangeByName("Project").getValue();
  var firstName = ss.getRangeByName("YourFirstName").getValue();
  var lastName = ss.getRangeByName("YourLastName").getValue();

  // Calculate the row and column numbers for the vendor email
  var rowNDA = parseInt(row) + ROWNDAOFFSET;
  var vendorEmail = sheet.getRange(rowNDA, COLUMNNDAOFFSET).getValue();

  // Display a toast message to show the NDA is being sent to the vendor
  ss.toast("🔨 Building NDA for " + vendorEmail, "Let's get this party started!", 20);

  // Call createPdfAndFolder function
  createPdfAndFolder(ss, sheet, rowNDA, COLUMNNDAOFFSET, projectName, firstName, lastName, vendorEmail);

  // Display a toast message to show the PDF has been placed in a folder and shared with the vendor
  ss.toast("The folder has been shared with " + vendorEmail + ".", "📂 PDF created.");

  // Log user activity as "NDA sent"
  logUserActivity("NDA Created");
}


/**
 * This function creates a folder containing a PDF generated from the "NDA TEMPLATE" sheet,
 * a copy of the PDF for the vendor to execute, and a README spreadsheet. The folder is then shared with the vendor.
 * It performs the following tasks:
 * Create the necessary folders for the project and NDA
 * Create an NDA template PDF if it doesn't exist
 * Create a PDF copy for the vendor
 * Create a README spreadsheet with instructions for the vendor
 * Share the folder with the vendor and update the folder URL in the sheet
 *
 * @param {object} ss - The current spreadsheet.
 * @param {object} sheet - The active sheet.
 * @param {number} row - The row number of the edited cell.
 * @param {number} column - The column number of the edited cell.
 * @param {string} projectName - The name of the project.
 * @param {string} firstName - The user's first name.
 * @param {string} lastName - The user's last name.
 * @param {string} vendorEmail - The email address of the vendor.
 * @param {string} ndaFileUrl - The URL of the existing NDA file (if any)
 * @returns {string} folderUrl - The URL of the created folder.
*/
function createPdfAndFolder(ss, sheet, row, column, projectName, firstName, lastName, vendorEmail) { 
  // Get the OpenBID folder
  var openBIDFolder = getOpenBIDFolder();

  // Create the folders  
  var projectFolder = findOrCreateFolder(openBIDFolder, projectName);
  var ndaFolder = findOrCreateFolder(projectFolder, "NDAs");

  // Create a folder with the specified name of the vendor and put it to the NDA folder
  var folderName = vendorEmail + "-NDA for " + projectName + " from " + firstName + " " + lastName;
  var folder = findOrCreateFolder(ndaFolder, folderName);

  // Create a NDA PDF for the vendor and place it in the folder
  createNDAForVendor(ss, vendorEmail, projectName, firstName, lastName, folder);

  // Share it with the vendor
  folder.addEditor(vendorEmail);

  // Create a README spreadsheet with instructions for the vendor and place it in the folder
  createReadmeSpreadsheet(folder, projectName, firstName, lastName);

  // Update the folder URL in the sheet and make the cell active
  var folderUrlCell = sheet.getRange(row, column + 1);
  folderUrlCell.setValue(folder.getUrl());
  folderUrlCell.activate();

  return folder.getUrl();
}


/**
 * Creates a PDF copy of an NDA file for a vendor, renames it, and adds it to a specified folder.
 *
 * @param {object} ss - The current spreadsheet.
 * @param {string} vendorEmail - The email address of the vendor to create the PDF copy for.
 * @param {string} projectName - The name of the project the NDA is for.
 * @param {string} firstName - The first name of the project owner.
 * @param {string} lastName - The last name of the project owner.
 * @param {Folder} folder - The folder to add the PDF copy to.
 */
function createNDAForVendor(ss, vendorEmail, projectName, firstName, lastName, folder) {
  // Generate a PDF from the "NDA TEMPLATE" sheet
  var ndaTemplateSheet = ss.getSheetByName("NDA TEMPLATE");
  var pdfFileName = vendorEmail + " to execute-NDA for " + projectName + " from " + firstName + " " + lastName;
  var pdfFile = generatePdfFromSheet(ss, ndaTemplateSheet, pdfFileName);

  // 'Move' the file to the new folder
  Drive.Files.update(
    { 'parents': [ { 'id': folder.getId() } ] }, 
    pdfFile.getId()
  );

}


/**
 * Creates a README spreadsheet in a specified folder and populates it with project information.
 *
 * @param {Folder} folder - The folder where the README spreadsheet should be created.
 * @param {string} projectName - The name of the project.
 * @param {string} firstName - The first name of the project owner.
 * @param {string} lastName - The last name of the project owner.
 */
function createReadmeSpreadsheet(folder, projectName, firstName, lastName) {
  // Create the README spreadsheet
  var readmeFileName = "README";
  var readmeSpreadsheet = SpreadsheetApp.create(readmeFileName);
  var readmeSpreadsheetId = readmeSpreadsheet.getId();

  // 'Move' the file to the new folder
  Drive.Files.update(
    { 'parents': [ { 'id': folder.getId() } ] }, 
    readmeSpreadsheetId
  );

  // Populate the README sheet with project information
  var readmeSheet = readmeSpreadsheet.getSheetByName("Sheet1");
  var text = firstName + " " + lastName + " would like you to bid on the project " + projectName + ". Please download the NDA PDF in this folder, sign with Adobe Fill & Sign, upload it back to this folder and email them when you're done.";
  readmeSheet.getRange("B2").setValue(text);
  readmeSheet.getRange("B2").setFontFamily("Lexend");

  // Hide gridlines
  readmeSheet.setHiddenGridlines(true);
}


/**
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 *
 * Send out BID functions
 * 
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 */


/**
 * This function creates and displays a modal dialog using an HTML file
 * named 'BidPopup'. The modal dialog is displayed within the Google Sheets
 * user interface, providing additional functionality for the OpenBID Toolkit.
 */
function showBidPopup() {
  // Log user activity as "Bid pop up shown"
  logUserActivity("Bid pop up shown");

  // Create an HTML output from the 'BidPopup' HTML file
  var htmlOutput = HtmlService.createHtmlOutputFromFile('BidPopup')
    .setWidth(900)
    .setHeight(550);

  // Display the modal dialog within the Google Sheets user interface
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'OpenBID Toolkit - Request Bids');
}


/**
 * Retrieves vendor bid data, episode list.
 *
 * @returns {Object} - An object containing the vendor list, episode list, or an error message.
 */
function getVendorBidData() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get the named ranges for vendor list and episode list
    var vendorListRange = ss.getRangeByName("vendorList");
    var episodeListRange = ss.getRangeByName("episodeList");

    // Retrieve the values from the named ranges
    var vendorList = vendorListRange.getValues();
    var episodeList = episodeListRange.getValues();

    // Return an object containing the vendor and episode data
    return { vendorList: vendorList, episodeList: episodeList };
  } catch (error) {
    // Log and return the error message
    // Logger.log('Error: ' + error.message);
    return {
      error: error.message
    };
  }
}


/**
 * This function sends a bid request to a vendor
 * and logs the activity for subscribed users.
 *
 * @param {number} column - The row number of the vendor in the "ADD VENDORS" sheet.
 * @param {number} row - The row number of the episode in the "OVERVIEW" sheet.
 * @param {string} vendorEmail - The email address in the row in the "ADD VENDORS" sheet.  
 */
function sendBid(row, column, vendorEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("OVERVIEW");

  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !userSubscriptions.includes("PaidSups")) {
    handleAccessDenied();
  } else {
    // User is a "PaidSups" subscriber or the subscription checking is not enabled.
    sendBidToVendor(ss, sheet, row, column, vendorEmail);
  }
  // Logger.log("DONE");
}


/**
 * This function sends a bid request to a vendor by creating a new Google Sheet for the bid and updating the appropriate cells in the original sheet.
 * It also logs the user's activity and highlights the relevant cell in the sheet.
 *
 * @param {Object} ss - The Spreadsheet object.
 * @param {Object} sheet - The active sheet within the spreadsheet.
 * @param {number} row - The row number where the vendor's email is located.
 * @param {number} column - The column number where the vendor's email is located.
 * @param {string} vendorEmail - The vendor's email address.
 */
function sendBidToVendor(ss, sheet, row, column, vendorEmail) {
  // Calculate the transposed row and column offsets for the bid
  var rowBid = column + ROWBIDOFFSET;
  var columnBid = row + COLUMNBIDOFFSET;

  // Inform the user that the script is working
  ss.toast("Checking if the bid already exists.", '⏱️ Stand By', 5);

  // Get the projectName name, episode, and package values from the row
  var projectName = sheet.getRange(rowBid, PROJECT_NAME_COLUMN).getValue();
  var episode = sheet.getRange(rowBid, EPISODE_COLUMN).getValue();
  var package = sheet.getRange(rowBid, PACKAGE_COLUMN).getValue();
  var vendorNumber = columnBid - VENDOR_EMAIL_OFFSET;
  // Logger.log('Project Name: ' + projectName + ', Episode: ' + episode + ', Package: ' + package + ', vendorNumber: ' + vendorNumber);

  // Construct the file name based on the values in the sheet Episode-Breakdown-email
  var fileName = vendorEmail + "-" + projectName + "-" + episode + "-Breakdown " + package;
  // Logger.log('File Name: ' + fileName);

  // Create the new files if needed
  var newFileURL = createOpenBIDSpreadsheet(fileName, projectName ,vendorNumber, vendorEmail, rowBid);
  // Logger.log('New File URL: ' + newFileURL);

  // Enter the new file URL
  sheet.getRange(rowBid, columnBid).setValue(newFileURL);

  // Get the last row with data in column B and set the newFileURL value
  var lastRowWithData = getLastRowWithData(sheet, ALLBIDURLS);
  var summaryURL = sheet.getRange(lastRowWithData + 1, ALLBIDURLS);
  summaryURL.setValue(newFileURL);
  // Logger.log('Last Row with Data: ' + lastRowWithData);

  // Highlight the cell with the url in column B
  summaryURL.activateAsCurrentCell();

  // Log activity
  logUserActivity("Bid sent");
}


/**
 * Main function for creating an OpenBID spreadsheet.
 *
 * @param {string} fileName - Name of the new spreadsheet file to be created.
 * @param {string} projectName - Name of the projectName folder where the new file should be created.
 * @param {number} vendorNumber - Row number of the vendor in the main sheet.
 * @param {string} vendorEmail - Email address of the vendor matching the vendor number.
 * @param {number} row - Row number of the selected episode.
 * @returns {string} URL of the created spreadsheet.
 */
function createOpenBIDSpreadsheet(fileName, projectName, vendorNumber, vendorEmail, row) {
  // Get the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Logger.log('Active spreadsheet obtained.');

  // Get the OpenBID folder
  var openBIDFolder = getOpenBIDFolder();

  // Create the folders  
  var projectFolder = findOrCreateFolder(openBIDFolder, projectName);

  // If a firewall file with the same name already exists in the projectFolder share it with the vendor and exit
  var existingFiles = projectFolder.getFilesByName(fileName);
  if (existingFiles.hasNext()) {

    // Tell them the file already exisits
    ss.toast("To start a new bidding round. Create a copy of this spreadsheet and increment the Breakdown Package number.", "⚠️️ OpenBID File For This Vendor Already Exists!", 40);

    // Get existing file URl
    var existingFirewall = existingFiles.next().getUrl();

    // Share it with vendorEmail
    existingFirewall.addViewer(vendorEmail);

    // Exit and return its URL
    return existingFirewall;
  }

  // Call the createFirewall function to create or find the "Firewall" spreadsheet
  var firewallSpreadsheetUrl = createFirewall(projectName);
 
  // Share the new firewall with the vendor
  var firewallFileId = firewallSpreadsheetUrl.match(/[-\w]{25,}/);
  var newFirewallFile = DriveApp.getFileById(firewallFileId);
  newFirewallFile.addViewer(vendorEmail);

  // Copy the template spreadsheet using the provided ID
  var template = DriveApp.getFileById(OPENBIDVENDOR_SHEET_ID);
  var newFile = template.makeCopy(fileName, projectFolder);

  // Tell them we're making the file  
  ss.toast("Creating " + fileName + ".", '🔨 Building...', 30);

  // Share the new bid with the vendorEmail
  newFile.addEditor(vendorEmail);

  // Open the new spreadsheet
  var newSpreadsheet = SpreadsheetApp.open(newFile);

  // Set the value of the named range "BidNumber" in the new spreadsheet
  var bidNumberSheetRange = newSpreadsheet.getRangeByName("BidNumber");
  if (bidNumberSheetRange) {
    bidNumberSheetRange.setFormula('=TRANSPOSE(IMPORTRANGE("' + firewallSpreadsheetUrl + '", "EPISODE INFO FOR VENDORS!D' + row + ':R' + row + '"))');
  } else {
    // If "BidNumber" named range does not exist, log a warning message
    // console.warn('Warning: "BidNumber" named range not found in the new spreadsheet.');
  }

  // Set the value of the named range "ClientSheet" in the new spreadsheet
  var clientSheetRange = newSpreadsheet.getRangeByName("ClientSheet");
  if (clientSheetRange) {
    clientSheetRange.setValue(firewallSpreadsheetUrl);
  } else {
    // If "ClientSheet" named range does not exist, log a warning message
    // console.warn('Warning: "ClientSheet" named range not found in the new spreadsheet.');
  }

  // Set the value of the named range "vendorNumber" in the new spreadsheet
  var vendorNumberRange = newSpreadsheet.getRangeByName("vendorNumber");
  if (vendorNumberRange) {
    vendorNumberRange.setValue(vendorNumber);
  } else {
    // If "vendorNumber" named range does not exist, log a warning message
    // console.warn('Warning: "vendorNumber" named range not found in the new spreadsheet.');
  }

  // Share the bid material link with the vendor
  shareBidMaterialLink(row,vendorEmail);

  // Tell them we sent the bid
  ss.toast("📤 Bid sent to " + vendorEmail + ".", "Booyaka!", 5);

  // Return the URL of the created spreadsheet
  return newFile.getUrl();
}


/**
 * This function creates a new "Firewall" spreadsheet containing SHOTS and ASSETS sheets.
 * These sheets are populated using IMPORTRANGE formulas that import data from a private spreadsheet.
 * The new "Firewall" spreadsheet is created inside the projectName folder within the "OpenBID" folder.
 * If a "Firewall" spreadsheet already exists for the projectName, the existing file's URL is returned.
 *
 * @param {string} projectName - Name of the projectName folder where the new "Firewall" file should be created.
 * @returns {string} URL of the created or existing "Firewall" spreadsheet.
 */
function createFirewall(projectName) {
  // Get the active spreadsheet and its name
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSpreadsheetName = ss.getName();
  var firewallFileName = "Firewall for " + activeSpreadsheetName;

  // Get the OpenBID folder
  var openBIDFolder = getOpenBIDFolder();

  // Create the folders  
  var projectFolder = findOrCreateFolder(openBIDFolder, projectName);

  // Check if a file with the same name already exists in the projectFolder
  var existingFiles = projectFolder.getFilesByName(firewallFileName);
  if (existingFiles.hasNext()) {
    // If the file already exists, return its URL
    ss.toast("Using " + firewallFileName + ".", "🔒 Better Safe Than Sorry!", 15);
    return existingFiles.next().getUrl();
  }

  // Create a new "Firewall" spreadsheet
  var newFirewall = SpreadsheetApp.create(firewallFileName);

  // Tell them we're making the file
  ss.toast("Using " + firewallFileName + ".", "🔒 Better Safe Than Sorry!", 50);

  // Move the new "Firewall" spreadsheet to the projectFolder
  var newFirewallFile = DriveApp.getFileById(newFirewall.getId());
  
  // 'Move' the file to the new folder
  Drive.Files.update(
    { 'parents': [ { 'id': projectFolder.getId() } ] }, 
    newFirewallFile.getId()
  );

  // Store the private spreadsheet URL
  var privateSpreadsheetUrl = ss.getUrl();

  // Create the EPISODE INFO FOR VENDORS sheet with named range and IMPORTRANGE formula
  var episodesSheet = newFirewall.insertSheet("EPISODE INFO FOR VENDORS");
  episodesSheet.getRange("D5").setFormula('=IMPORTRANGE("' + privateSpreadsheetUrl + '", "FromClientAllEpisodes")');
  // Logger.log("EPISODE INFO FOR VENDORS sheet created with IMPORTRANGE formula.");

  // Add some instructions
  episodesSheet.getRange("D4").setValue("⬇ Allow access below, then ignore this sheet.");

  // Create the SHOTS sheet with named range and IMPORTRANGE formula
  var shotsSheet = newFirewall.insertSheet("SHOTS");
  var shotsNamedRange = "FromClientAllShots";
  newFirewall.setNamedRange(shotsNamedRange, shotsSheet.getRange("B7:AK506"));
  shotsSheet.getRange("B7").setFormula('=IMPORTRANGE("' + privateSpreadsheetUrl + '", "' + shotsNamedRange + '")');
  // Logger.log("SHOTS sheet created with named range and IMPORTRANGE formula.");

  // Create the ASSETS sheet with named range and IMPORTRANGE formula
  var assetsSheet = newFirewall.insertSheet("ASSETS");
  var assetsNamedRange = "FromClientAllAssets";
  newFirewall.setNamedRange(assetsNamedRange, assetsSheet.getRange("B7:AF56"));
  assetsSheet.getRange("B7").setFormula('=IMPORTRANGE("' + privateSpreadsheetUrl + '", "' + assetsNamedRange + '")');
  // Logger.log("ASSETS sheet created with named range and IMPORTRANGE formula.");

  // Delete the default "Sheet1" from the newFirewall spreadsheet
  var sheetToDelete = newFirewall.getSheetByName("Sheet1");
  newFirewall.deleteSheet(sheetToDelete);

  // Get the firewallFile named range
  var firewallFileRange = ss.getRangeByName("firewallFile");

  // Get the newFirewall Url
  var newFirewallUrl = newFirewall.getUrl();

  // Display a popup with a red button linked to the firewallSpreadsheetUrl
  showFirewallPopup(newFirewallUrl);

  // Put the url in the cel
  firewallFileRange.setValue(newFirewallUrl);

  return newFirewall.getUrl();
}


/**
 * Shares the bid material link with the specified vendor email address.
 *
 * @param {number} row - The row number of the bid material link in the active sheet.
 * @param {string} vendorEmail - The email address of the vendor to share the bid material link with.
 */
function shareBidMaterialLink(row, vendorEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("OVERVIEW");
 
  // Get the bid material link from the sheet
  var bidMaterialLink = sheet.getRange(row, BID_MATERIAL).getValue();
 
  // If bidMaterialLink is empty or not a Google Drive link, log a warning and exit
  if (!bidMaterialLink || !bidMaterialLink.startsWith("https://drive.google.com/")) {
    ss.toast("The provided bid material link is empty or not a valid Google Drive link.","⚠️️ Warning");
    return;
  }
 
  // Get the file ID from the Google Drive link
  var fileId = bidMaterialLink.match(/[-\w]{25,}/);
 
  // If file ID is empty or invalid, log a warning and exit
  if (!fileId || fileId.length < 1) {
    ss.toast("Failed to extract file ID from the Google Drive bid material link.","⚠️️ Warning");
    return;
  }
 
  try {
    // Get the file from Google Drive
    var bidMaterialFile = DriveApp.getFileById(fileId[0]);
 
    // Share the file with the specified vendor email address
    bidMaterialFile.addViewer(vendorEmail);
 
    // Log the success message
    ss.toast("Successfully shared the bid material link with " + vendorEmail, "👍 Great");
  } catch (error) {
    // Log the error message
    ss.toast("⚠️️ Error sharing bid material link with " + vendorEmail + ": " + error);
  }
}


/**
 * Displays a popup to inform the user about the created firewall spreadsheet.
 * The popup contains a red button that links to the firewall spreadsheet URL.
 *
 * @param {string} firewallSpreadsheetUrl - URL of the created firewall spreadsheet.
 */
function showFirewallPopup(firewallSpreadsheetUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Create a popup with a red button linked to the firewallSpreadsheetUrl using the HTML file
  var htmlTemplate = HtmlService.createTemplateFromFile('firewallPopup');
  htmlTemplate.firewallSpreadsheetUrl = firewallSpreadsheetUrl;
  var htmlOutput = htmlTemplate.evaluate().setWidth(400).setHeight(450);

  // Show the popup
  ui.showModalDialog(htmlOutput, '100% secure!');
}


/**
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 *
 * Insert Thumbnail functions
 * 
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 */


/**
 * Displays a sidebar in the active Google Sheet with a user interface for
 * importing images from a Google Drive folder into the sheet. Users can enter
 * the folder URL, and the images will be displayed along with buttons for
 * inserting the images into the selected cells.
 */
function showImageSidebar() {
  // Log user activity
  logUserActivity("Image sidebar opened");

    // Create an HTML output from the 'sidebar' file
  var html = HtmlService.createHtmlOutputFromFile('imageSidebar')
      .setTitle('Thumbnail Importer') // Set the title of the sidebar
      .setWidth(300); // Set the width of the sidebar

  // Display the sidebar in the active Google Sheet
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Retrieves images from the specified Google Drive folder.
 *
 * @param {string} folderUrl - The URL of the Google Drive folder.
 * @returns {Array} images - An array of objects containing the image name and URL.
 */
function getImages(folderUrl) {
  var folderId = folderUrl.match(/[-\w]{25,}/);
  var folder = DriveApp.getFolderById(folderId[0]);

  // Check the access level of the folder
  var access = folder.getAccess(DriveApp.Access.ANYONE_WITH_LINK);
  if (access != DriveApp.Permission.VIEW) {
    throw new Error("The =IMAGE() function requires the folder permissions be set to 'Anyone with the link'. If this is unsuitable for you, you must manually insert images into cells using the Insert menu.");
  }

  var files = folder.getFiles();
  var images = [];
  var imageMimeTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/bmp', 'image/svg+xml'];

  while (files.hasNext()) {
    var file = files.next();
   
    if (imageMimeTypes.includes(file.getMimeType())) {
      var directImageUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();
      images.push({
        name: file.getName(),
        url: directImageUrl
      });
    }
  }

  // Log user activity
  logUserActivity("Got images")

  return images;
}


/**
 * Inserts an image into the selected cell using the IMAGE formula.
 *
 * @param {string} imageUrl - The URL of the image to be inserted.
 */
function insertImage(imageUrl) {
  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" or "PaidUsers" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !(userSubscriptions.includes("PaidSups") || userSubscriptions.includes("PaidUsers"))) {
    handleAccessDenied();
  } else {
    // User is a "PaidSups" or "PaidUsers" subscriber or the subscription checking is not enabled.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var selectedRange = sheet.getActiveRange();
    var selectedRow = selectedRange.getRow();
    var selectedColumn = selectedRange.getColumn();

    var imageFormula = `=image("${imageUrl}")`;
    sheet.getRange(selectedRow, selectedColumn).setFormula(imageFormula);

    // Log user activity
    logUserActivity("Inserted an image");
  }
}


/**
 * Automatically inserts images that match the IDs in the named ranges.
 *
 * @param {string} folderUrl - The URL of the Google Drive folder.
 * @returns {Object} result - An object containing the status message or error message.
 */
function autoInsertImages(folderUrl) {
  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" or "PaidUsers" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !(userSubscriptions.includes("PaidSups") || userSubscriptions.includes("PaidUsers"))) {
    handleAccessDenied();
  } else {
    // User is a "PaidSups" or "PaidUsers" subscriber or the subscription checking is not enabled.
    var result = {};

    // Check if the folder URL is provided
    if (!folderUrl || folderUrl.trim() === '') {
      result.error = "No folder URL provided";
      return result;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    // Define the named ranges to search for
    var namedRanges = ["shotIDs", "assetIDs"];

    var namedRangesFound = false;

    // Fetch the images from the provided folder URL
    var images = getImages(folderUrl);

    // Prepare an object to store the formulas to be inserted
    var formulasToInsert = {};

    // Iterate through the named ranges
    namedRanges.forEach(function (namedRange) {
      // Check if the named range exists in the active sheet
      var range = ss.getRangeByName(namedRange);
      if (range && range.getSheet().getName() === sheet.getName()) {
        namedRangesFound = true;
        var values = range.getValues();

        // Iterate through the fetched images
        images.forEach(function (image) {
          var imageName = image.name.split('.').slice(0, -1).join('.'); // Remove the file extension

          // Compare image names with the values in the named range
          for (var i = 0; i < values.length; i++) {
            if (values[i][0] === imageName) {
              // If a match is found, store the formula to be inserted in the corresponding row
              var imageFormula = `=image("${image.url}")`;
              formulasToInsert[range.getRow() + i] = imageFormula;
            }
          }
        });
      }
    });

    // If no named ranges are found, set the error message
    if (!namedRangesFound) {
      result.error = "No IDs found";
    } else {
      // If named ranges are found, insert the image formulas in one go
      for (var row in formulasToInsert) {
        sheet.getRange(row, 3).setFormula(formulasToInsert[row]);
      }
      result.message = "Images inserted successfully";
    }

    // Log user activity
    logUserActivity("Auto-inserted images");

    return result;
  }
}


/**
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 *
 * YouTube Chapter functions
 * 
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 */


/**
 * Displays a sidebar in the active Google Sheet with a user interface for
 * importing chapters from a YouTube video into the sheet. Users can enter
 * the YouTube URL, and the chapters will be displayed along with buttons for
 * inserting the links into the selected cells.
 */
function showChapterSidebar() {
  // Log user activity
  logUserActivity("Chapter sidebar opened");

  // Create an HTML output from the 'chapterSidebar' file
  var html = HtmlService.createHtmlOutputFromFile('chapterSidebar')
      .setTitle('YouTube Chapter Importer') // Set the title of the sidebar
      .setWidth(300); // Set the width of the sidebar

  // Display the sidebar in the active Google Sheet
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Extracts chapters from the description of a YouTube video.
 *
 * @param {string} videoUrl - The URL of the YouTube video.
 * @returns {Array} - An array of chapter objects, each containing a name, time, and YouTube URL.
 */
function extractChapters(videoUrl) {
  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // Extract the video ID from the provided YouTube URL.
  var videoId = extractVideoIdFromUrl(videoUrl);
  if (!videoId) {
    throw new Error('Invalid YouTube URL');
  }

  // Retrieve the video description using the video ID.
  var description = getPrivateVideoDescription(videoId);
  
  // Parse the chapters from the video description.
  var chapters = parseChapters(description, videoId);

  // Log user activity
  logUserActivity("Got chapters");

  // Return the extracted chapters.
  return chapters;
  
}


/**
 * Parses the chapters from the given video description.
 *
 * @param {string} description - The video description containing chapter information.
 * @returns {Array} - An array of chapter objects, each containing a name, time, and YouTube URL.
 */
function parseChapters(description, videoId) {
  var chapters = [];
  var lines = description.split('\n');
  var chapterRegex = /^((?:\d+:)?\d+:\d+)\s+(.+)$/;

  lines.forEach(function (line) {
    var match = line.match(chapterRegex);
    if (match) {
      var time = match[1];
      var name = match[2].trim();

      // Remove "- " from the beginning of the chapter name, if needed
      if (name.startsWith("- ")) {
        name = name.substring(2);
      }
      var seconds = timeStringToSeconds(time);
      chapters.push({
        name: name,
        time: seconds,
        youtubeUrl: 'https://youtu.be/' + videoId,
      });
    }
  });

  // Return the result object.
  return chapters;
}


/**
 * Converts a time string (hh:mm:ss or mm:ss) to seconds.
 *
 * @param {string} timeString - The time string to convert.
 * @returns {number} The time in seconds.
 */
function timeStringToSeconds(timeString) {
  var parts = timeString.split(':');
  var seconds = 0;
  for (var i = 0; i < parts.length; i++) {
    seconds = seconds * 60 + parseInt(parts[i], 10);
  }

  // Return the result object.
  return seconds;
}


/**
 * Automatically inserts chapters as hyperlinks into a Google Sheet.
 *
 * @param {string} youtubeUrl - The URL of the YouTube video containing chapters.
 * @returns {Object} - An object containing a message or error describing the result of the operation.
 */
function autoInsertChapters(youtubeUrl) {
  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" or "PaidUsers" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !(userSubscriptions.includes("PaidSups") || userSubscriptions.includes("PaidUsers"))) {
    handleAccessDenied();
  } else {
    // User is a "PaidSups" or "PaidUsers" subscriber or the subscription checking is not enabled.
    // Initialize the result object.
    var result = {};

    // Check if a valid YouTube URL is provided.
    if (!youtubeUrl || youtubeUrl.trim() === "") {
      result.error = "No YouTube URL provided";
      return result;
    }

    // Get the active Google Sheet and its active sheet.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    // Define the named ranges to look for.
    var namedRanges = ["shotIDs"];
    var namedRangesFound = false;

    // Extract chapters from the YouTube URL.
    var chapters = extractChapters(youtubeUrl);

    // If no chapters were found, return an error.
    if (chapters.length === 0) {
      result.error = "No chapters found in the provided YouTube URL";
      return result;
    }
    
    var formulasToInsert = {};

    // Iterate through the named ranges.
    namedRanges.forEach(function (namedRange) {
      var range = ss.getRangeByName(namedRange);
      // Check if the range exists and is in the current sheet.
      if (range && range.getSheet().getName() === sheet.getName()) {
        namedRangesFound = true;
        var values = range.getValues();

        // Iterate through the chapters and find matching chapter names in the range.
        chapters.forEach(function (chapter) {
          for (var i = 0; i < values.length; i++) {
            if (values[i][0] === chapter.name) {
              // Create a hyperlink formula for the chapter.
              var hyperlinkFormula = `=hyperlink("${chapter.youtubeUrl}?t=${chapter.time}", "${chapter.name}")`;
              // Add the formula to the formulasToInsert object.
              formulasToInsert[range.getRow() + i] = hyperlinkFormula;
            }
          }
        });
      }
    });

    // If no named ranges were found, return an error.
    if (!namedRangesFound) {
      result.error = "No IDs found";
    } else {
      // Insert the hyperlink formulas into the sheet.
      for (var row in formulasToInsert) {
        sheet.getRange(row, PUTYOUTUBELINK).setFormula(formulasToInsert[row]);
      }
      // Set the success message.
      result.message = "Chapters inserted successfully";
    }

    // Log user activity
    logUserActivity("Auto-inserted chapters");

    // Return the result object.
    return result;
  }
}


/**
 * Inserts a single chapter as a hyperlink in the active cell of the Google Sheet.
 *
 * @param {Object} chapter - An object containing the chapter's name, YouTube URL, and time.
 */
function insertChapter(chapter) {
  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" or "PaidUsers" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !(userSubscriptions.includes("PaidSups") || userSubscriptions.includes("PaidUsers"))) {
    handleAccessDenied();
  } else {
    // User is a "PaidSups" or "PaidUsers" subscriber or the subscription checking is not enabled.
    
    // Get the active Google Sheet and its active sheet.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    // Get the active range, row, and column in the sheet.
    var selectedRange = sheet.getActiveRange();
    var selectedRow = selectedRange.getRow();
    var selectedColumn = selectedRange.getColumn();

    // Create a hyperlink formula for the chapter.
    var hyperlinkFormula = `=hyperlink("${chapter.youtubeUrl}?t=${chapter.time}", "${chapter.name}")`;

    // Set the formula in the active cell.
    sheet.getRange(selectedRow, selectedColumn).setFormula(hyperlinkFormula);

    // Log user activity
    logUserActivity("Inserted a chapter");
  }
}


/**
 * Retrieves the description of a private YouTube video using the YouTube API.
 *
 * @param {string} videoId - The ID of the YouTube video.
 * @returns {string} The video's description.
 * @throws {Error} If the video is not found or not accessible.
 */
function getPrivateVideoDescription(videoId) {
  // Call the YouTube API to get the video snippet for the specified video ID.
  var video = YouTube.Videos.list('snippet', { id: videoId });

  // If the video is found, return its description.
  if (video.items.length > 0) {
    return video.items[0].snippet.description;
  } else {
    // If the video is not found or not accessible, throw an error.
    throw new Error('Video not found or not accessible.');
  }
}


/**
 * Extracts the YouTube video ID from a given URL.
 *
 * @param {string} url - The YouTube video URL.
 * @returns {string|null} The extracted video ID, or null if the URL is not valid.
 */
function extractVideoIdFromUrl(url) {
  // Define the regular expression pattern to match a YouTube URL and capture the video ID.
  var regex = /^.*(youtu.be\/|v\/|u\/\w\/|embed\/|watch\?v=|\&v=)([^#\&\?]*).*/;

  // Test the URL against the regular expression pattern.
  var match = url.match(regex);

  // If there's a match and the video ID is 11 characters long, return the video ID.
  if (match && match[2].length === 11) {
    return match[2];
  } else {
    // If the URL is not valid, return null.
    return null;
  }
}


/**
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 *
 * Subscription functions
 * 
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 */


/**
 * This function wraps the checkUserSubscriptionStatus function and handles the case
 * when a user is not subscribed to the service. If the subscription check is disabled,
 * it will always return ['No Subscription'].
 *
 * @return {Array} - Returns an array containing the subscription types of the user if the check is enabled,
 *                     ['No Subscription'] otherwise.
 * 
 */
function isUserSubscribed() {
  if (ENABLE_SUBSCRIPTION_CHECK) {
    return checkUserSubscriptionStatus();
  } else {
    return ["No Subscription"];
  }
}


/**
 * Function to check if the user's email is in the list of authorised subscribers.
 *
 * @return {Array} - Returns an array containing 'PaidUsers', 'PaidSups', and/or 'PaidSupport' based on the user's subscription levels.
 */
function checkUserSubscriptionStatus() {
  // Check to see if the user has subscribed
  var userEmail = Session.getEffectiveUser().getEmail();

  // Access the Google Sheet containing the 'PaidUsers', 'PaidSups', 'PaidSupport' named ranges using its ID
  var subscriptionsSheet = SpreadsheetApp.openById(SUBSCRIBER_SHEET_ID);

  // Get the values in the 'PaidUsers', 'PaidSups', 'PaidSupport' named ranges
  var paidUsersRange = subscriptionsSheet.getRange('PaidUsers');
  var paidSupsRange = subscriptionsSheet.getRange('PaidSups');
  var paidSupportRange = subscriptionsSheet.getRange('PaidSupport');

  var paidUsers = paidUsersRange.getValues();
  var paidSups = paidSupsRange.getValues();
  var paidSupport = paidSupportRange.getValues();

  var userSubscriptions = [];

  // Check if the user's email is in any of the lists of authorised users
  if(isUserInList(userEmail, paidUsers)) {
    userSubscriptions.push("PaidUsers");
  }
  if(isUserInList(userEmail, paidSups)) {
    userSubscriptions.push("PaidSups");
  }
  if(isUserInList(userEmail, paidSupport)) {
    userSubscriptions.push("PaidSupport");
  }

  if(userSubscriptions.length == 0) {
    userSubscriptions.push("No Subscription");
  }

  // Return the result object.
  return userSubscriptions;
}


/**
 * Checks if the given user email is in the list of paid users.
 *
 * @param {string} userEmail - The email of the user to check.
 * @param {Array} paidUsers - A 2D array containing the list of paid users' emails.
 * @return {boolean} - True if the userEmail is in the list, false otherwise.
 */
function isUserInList(userEmail, paidUsers) {
  for (var i = 0; i < paidUsers.length; i++) {
    // Logger.log('Checking against: ' + paidUsers[i][0]);
    if (paidUsers[i][0] === userEmail) {
      // Logger.log(userEmail + ' is on the list');
      return true;
    }
  }
  return false;
}


/**
 * This function handles the case when a user is not subscribed to the service.
 * It logs the user's activity as "Access denied" and displays a popup message.
 */
function handleAccessDenied() {
  // Log the user's subscription status
  // Logger.log('User is not subscribed');

  // Call the accessDenied function to display a popup message
  accessDenied();

  // Log user activity as "Access denied"
  logUserActivity("Access denied");
}


/**
 * Function to open the Access Denied dialogue box.
 *
 */
function accessDenied() {
  // Show a custom modal dialog with a clickable link if the user is not in the list of authorised users
  var htmlOutput = HtmlService.createHtmlOutputFromFile('accessDeniedDialog.html')
    .setWidth(680)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'About Your Free Trial');
}


/**
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 *
 * Vendor functions
 * 
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 */


/**
 * This script copies data between the "SHOT BID BREAKDOWN" sheet and the "SHOT TEMPLATES" sheet
 * within the same Google Spreadsheet document. It operates on the currently selected range in 
 * the "SHOT BID BREAKDOWN" sheet, copying certain values to the "SHOT TEMPLATES" sheet, and then
 * pastes formulas from the last row of the "SHOT BID BREAKDOWN" sheet into the initial active range.
 */
function createTemplateShot() {
  // Log user activity
  logUserActivity("Creating a shot template");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();

  ss.toast("Creating a template from the selected shot.", "Working 🔨", 20);

  // Check we have all the data needed
  
  // Ensure the active sheet is "SHOT BID BREAKDOWN"
  if (activeSheet.getName() !== "SHOT BID BREAKDOWN") {
    var errorMessage = "The active sheet is not 'SHOT BID BREAKDOWN'";
    ss.toast(errorMessage,"Error ⚠️");
    return;
  }
  
  // Check if the cell in column F of the topRow is not empty
  var activeRange = activeSheet.getActiveRange();
  var topRow = activeRange.getRow();

  var cellFValue = activeSheet.getRange(topRow, 6).getValue();
  if (cellFValue !== "") {
    var errorMessage = "This shot is already a template";
    ss.toast(errorMessage, "Error ⚠️");
    return;
  }

  // Check if both name and difficulty values are blank
  var colHValue = activeSheet.getRange(topRow, 8).getValue();
  var colIValue = activeSheet.getRange(topRow, 9).getValue();
  
  if (colHValue === "" && colIValue === "") {
    var errorMessage = "Both columns H and I are blank. At least one should have a value.";
    ss.toast(errorMessage, "Error ⚠️");
    return;
  }

  // Switch to the SHOT TEMPLATES sheet
  var templatesSheet = ss.getSheetByName("SHOT TEMPLATES");
  if (!templatesSheet) {
    var errorMessage = "Error: The 'SHOT TEMPLATES' sheet does not exist";
    ss.toast(errorMessage, "Error ⚠️");
    return;
  }
  
  // Get the max row from column B in the SHOT BID BREAKDOWN sheet
  var maxRowB = activeSheet.getLastRow();

  // Copy the formulas from columns Q - AA and AE - AU
  var formulasFromQToAA = activeSheet.getRange(maxRowB, 17, 1, 11).getFormulas()[0];
  var formulasFromAEToAU = activeSheet.getRange(maxRowB, 31, 1, 17).getFormulas()[0];

  // Check if any formulas are missing from the template row
  for (var i = 0; i < formulasFromQToAA.length; i++) {
    if (!formulasFromQToAA[i]) {
        var errorMessage = "Your template row at the bottom is missing a formula!";
        ss.toast(errorMessage, "Error ⚠️");
        return;
    }
  }

  for (var i = 0; i < formulasFromAEToAU.length; i++) {
    if (!formulasFromAEToAU[i]) {
        var errorMessage = "Your template row at the bottom of this tab is missing a formula!";
        ss.toast(errorMessage, "Error ⚠️");
        return;
    }
  }

  // Checks - End

  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !userSubscriptions.includes("PaidUsers")) {
    handleAccessDenied();
  } else {
    // User is a "PaidUsers" subscriber or the subscription checking is not enabled.

    // Get values from columns Q to AT from the bid breakdown sheet
    var rangeValues = activeSheet.getRange(topRow, 17, 1, 30).getValues();

    // Get formulas from columns Q to AT from the bid breakdown sheet
    var rangeFormulas = activeSheet.getRange(topRow, 17, 1, 30).getFormulas();

    // Get the last row with content in column D of the template sheet
    var lastRowD = getLastRowWithContent(templatesSheet, "D");

    // Get the last row with content in column E of the template sheet
    var lastRowE = getLastRowWithContent(templatesSheet, "E");

    // Select the maximum row between columns D and E and add 1
    var lastRowDE = Math.max(lastRowD, lastRowE) + 1;

    // Paste the name and difficulty values into column D and E of the SHOT TEMPLATES sheet
    templatesSheet.getRange(lastRowDE, 4).setValue(colHValue);
    templatesSheet.getRange(lastRowDE, 5).setValue(colIValue);

    // Paste the values/formulas from columns Q to AT in SHOT BID BREAKDOWN sheet into columns G - Q and S - AH in SHOT TEMPLATES sheet
    var targetColumn = 7; // Starting with column G
    for (var sourceColumn = 0; sourceColumn <= 29; sourceColumn++) {
        if(sourceColumn !== 11 && sourceColumn !== 12 && sourceColumn !== 13) { // Exclude source columns 28, 29, 30
            var cellFormula = rangeFormulas[0][sourceColumn];
            // If the formula contains a sheet reference it will work in the template sheet. Assume a sheet reference is included in the formula if an exclamation mark '!' is found.
            if (cellFormula !== "" && cellFormula.includes('!')) { 
                templatesSheet.getRange(lastRowDE, targetColumn).setFormula(cellFormula);
            } else { 
                templatesSheet.getRange(lastRowDE, targetColumn).setValue(rangeValues[0][sourceColumn]);
            }

            // Increase target column (G - Q and then S - AH)
            if(targetColumn === 17) { // If reached end of Q, jump to S
                targetColumn = 19;
            } else {
                targetColumn++;
            }
        }
    }

    // Adjust the formulas to match the row being pasted into
    var adjustedFormulasFromQToAA = adjustFormulasRow(formulasFromQToAA, topRow);
    var adjustedFormulasFromAEToAU = adjustFormulasRow(formulasFromAEToAU, topRow);

    // Create a new 2D array to hold the formulas for the range
    var newFormulasFromQToAA = [adjustedFormulasFromQToAA];
    var newFormulasFromAEToAU = [adjustedFormulasFromAEToAU];

    // Use setFormulas to apply the new formulas to the entire range at once
    activeSheet.getRange(topRow, 17, 1, adjustedFormulasFromQToAA.length).setFormulas(newFormulasFromQToAA);
    activeSheet.getRange(topRow, 31, 1, adjustedFormulasFromAEToAU.length).setFormulas(newFormulasFromAEToAU);

    // Set the formula of column F of the top row to reference column F of the last row in the SHOT TEMPLATES sheet
    activeSheet.getRange(topRow, 6).setFormula("='SHOT TEMPLATES'!$F$" + lastRowDE);

    // Set the active cell to the template column
    activeSheet.setActiveRange(activeSheet.getRange("G" + topRow));

    ss.toast("Template created.", "👍 Nice!");
  }
}


/**
 * This script hides certain sheets and columns within a Google Spreadsheet.
 * It targets all sheets except for "SUMMARY", "GRAND TOTALS", "ASSET BID BREAKDOWN", and "SHOT BID BREAKDOWN", hiding them from view.
 * Within the "ASSET BID BREAKDOWN" and "SHOT BID BREAKDOWN" sheets, it hides specific columns for a cleaner, more focused view.
 * The targeted columns are columns O to BD for "ASSET BID BREAKDOWN" and columns F to G and AD to CL for "SHOT BID BREAKDOWN".
 *
 * @param {Object} ss - A Google Spreadsheet object
 */
function prepBid(ss) {
  // Log user activity
  logUserActivity("Preparing bid to share");

  // Check we are not in OpenBID Client or Vendor Template spreadsheets and stop execution if it returns false
  if (!checkInOpenBidVendor()) {
    return;
  }
  
  // Set the default to be the active sheet
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  ss.toast("Hiding everything sensitive!", "Working 🔨", 20);

  // Get all sheets in the Spreadsheet
  var allSheets = ss.getSheets();
  
  // List of sheet names that should remain visible
  var visibleSheetNames = ["SUMMARY", "GRAND TOTALS", "ASSET BID BREAKDOWN", "SHOT BID BREAKDOWN"];

  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !userSubscriptions.includes("PaidUsers")) {
    handleAccessDenied();
  } else {
    // User is a "PaidUsers" subscriber or the subscription checking is not enabled.

    try {
      // Loop through all sheets and hide any that aren't in the visibleSheetNames list
      for (var i = 0; i < allSheets.length; i++) {
        if (visibleSheetNames.indexOf(allSheets[i].getName()) == -1) {
          allSheets[i].hideSheet();
        }
      }
    } catch (error) {
      ss.toast("You can only do this in an OpenBID spreadsheet.", "⚠️ Error");
      // Logger.log('An error occurred: ' + error.toString());
    }


    // Retrieve the "ASSET BID BREAKDOWN" sheet and hide columns  O to BD (15 to 56)
    var assetSheet = ss.getSheetByName("ASSET BID BREAKDOWN");
    assetSheet.hideColumns(15, 42);
    // Logger.log("Hiding rows for ASSET BID BREAKDOWN");
    hideEmptyRows(assetSheet);
    
    // Retrieve the "SHOT BID BREAKDOWN" sheet and hide columns F to G and AE to CL (6 to 7 and 31 to 91)
    var shotSheet = ss.getSheetByName("SHOT BID BREAKDOWN");
    shotSheet.hideColumns(6, 2);  // Hide columns F to G
    shotSheet.hideColumns(31, 60); // Hide columns AE to CL
    // Logger.log("Hiding rows for SHOT BID BREAKDOWN");
    hideEmptyRows(shotSheet);

    ss.toast("Now share this spreadsheet with them with 'Comment Only' permissions.", "👍 Prep for client done!");
  }
}


/**
 * This function makes all sheets in the Google Spreadsheet visible again (except "#TAGS"),
 * unhides previously hidden columns in the "ASSET BID BREAKDOWN" and "SHOT BID BREAKDOWN" sheets,
 * and also unhides all previously hidden rows.
 *
 * @param {Object} [ss] - An optional Spreadsheet instance. If not provided, the active spreadsheet will be used.
 */
function undoPrepBid(ss) {
  // Log user activity
  logUserActivity("Undoing Prepped bid");

  // Check we are not in OpenBID Client or Vendor Template spreadsheets and stop execution if it returns false
  if (!checkInOpenBidVendor()) {
    return;
  }

  // If no Spreadsheet instance was provided, get the active one
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  ss.toast("Unhiding everything sensitive!", "Working 🔨", 20);

  // Get all sheets in the Spreadsheet
  var allSheets = ss.getSheets();

  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !userSubscriptions.includes("PaidUsers")) {
    handleAccessDenied();
  } else {
    // User is a "PaidUsers" subscriber or the subscription checking is not enabled.

    // Loop through all sheets
    for (var i = 0; i < allSheets.length; i++) {
      // Check if the sheet is not "#TAGS", "EXPORT ASSETS" or "EXPORT SHOTS"
      if (allSheets[i].getName() !== "#TAGS" && allSheets[i].getName() !== "EXPORT ASSETS" && allSheets[i].getName() !== "EXPORT SHOTS") {
        // If it's hidden, show the sheet
        if (allSheets[i].isSheetHidden()) {
          allSheets[i].showSheet();
        }
      }
    }

    // Retrieve the "ASSET BID BREAKDOWN" sheet and unhide columns O to BD (15 to 56)
    var assetSheet = ss.getSheetByName("ASSET BID BREAKDOWN");
    assetSheet.unhideColumn(assetSheet.getRange("O1:BD1"));

    // Unhide all rows in the "ASSET BID BREAKDOWN" sheet
    assetSheet.showRows(1, assetSheet.getMaxRows());

    // Retrieve the "SHOT BID BREAKDOWN" sheet and unhide columns F to G and AE to CL (6 to 7 and 31 to 91)
    var shotSheet = ss.getSheetByName("SHOT BID BREAKDOWN");
    shotSheet.unhideColumn(shotSheet.getRange("F1:G1"));
    shotSheet.unhideColumn(shotSheet.getRange("AE1:CL1"));

    // Unhide all rows in the "SHOT BID BREAKDOWN" sheet
    shotSheet.showRows(1, shotSheet.getMaxRows());

    ss.toast("Prep for client undone!", "Looks Good 👍");
  }
}


/**
 * This function creates a backup copy of the spreadsheet then in the new back up spreadsheet converts columns R to W to checkboxes with custom checked and unchecked values in the 
 * "SHOT BID BREAKDOWN" sheet and then gives the user a pop up to go to the file and download it.
 */
function bidExportPdf() {
  // Log user activity
  logUserActivity("Exporting a PDF bid");

  // Check we are not in OpenBID Client or Vendor Template spreadsheets and stop execution if it returns false
  if (!checkInOpenBidVendor()) {
    return;
  }

  // Retrieve the active Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Building a backup spreadsheet to export from.", "Working 🔨 ", 20);

  // Retrieve the named range "BidNumber"
  var namedRange = ss.getRangeByName("BidNumber");
  if (!namedRange) {
    ss.toast("Named range BidNumber not found", "⚠️ Error");
    throw new Error('Named range "BidNumber" not found');
  }

  // Retrieve the value of the named range
  var bidNumber = namedRange.getValue();
  var number = typeof bidNumber === 'string' ? parseInt(bidNumber.match(/\d+/)[0], 10) : bidNumber;

  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !userSubscriptions.includes("PaidUsers")) {
    handleAccessDenied();
  } else {
    // User is a "PaidUsers" subscriber or the subscription checking is not enabled.

    // Create the new file name
    var backupSpreadsheetName = ss.getName() + "Backup of Bid " + (isNaN(number) ? 0 : number);

    // Make a copy of the spreadsheet with the old name + old bid number
    var backupSpreadsheetFile = Drive.Files.copy({title: backupSpreadsheetName}, ss.getId());

    // Open the copy as a Spreadsheet object
    var backupSpreadsheet = SpreadsheetApp.openById(backupSpreadsheetFile.id);

    // Increment the bid number
    var newNumber = number + 1;

    // Get its URL and give it to the user
    var backupSpreadsheetURL = backupSpreadsheet.getUrl();
    exportPDFPopup(backupSpreadsheetURL, newNumber);

    // Prep the new spreadsheet by hiding stuff
    prepBid(backupSpreadsheet);

    // Convert checkboxes
    // Retrieve the "SHOT BID BREAKDOWN" sheet
    var shotSheet = backupSpreadsheet.getSheetByName("SHOT BID BREAKDOWN");

    // Define the range for columns R6 to W (18 to 23)
    var range = shotSheet.getRange("R6:W" + shotSheet.getLastRow());

    // Apply data validation rule to the defined range that requires the cell to be a checkbox with custom checked and unchecked values
    range.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox('✅', '⬜').build());

    // Update the client summary sheet and bid number in the master sheet
    updateClientSummary(ss, newNumber);
  
  }
}



/**
 * This function retrieves the value of the named range "BidNumber", extracts the number from the string, 
 * makes a copy of the spreadsheet with a new name, and updates the "BidNumber" value 
 * and paste the values and formatting of the "ClientSummary" range in the row immediately after the last 
 * row with data in the "SUMMARY" sheet of the new spreadsheet.
 *
 * @param {string} backupSpreadsheetName - Name for the backup spreadsheet
 */
function createNewBid(backupSpreadsheetName) {
  // Log user activity
  logUserActivity("Creating a bid version");

  // Check we are not in OpenBID Client or Vendor Template spreadsheets and stop execution if it returns false
  if (!checkInOpenBidVendor()) {
    return;
  }
  

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Retrieve the named range "BidNumber"
  var namedRange = ss.getRangeByName("BidNumber");
  if (!namedRange) {
    ss.toast("You can only do this in an OpenBID spreadsheet.", "⚠️ Error");
    throw new Error('Named range "BidNumber" not found');
  }

  // Retrieve the value of the named range
  var bidNumber = namedRange.getValue();

  // Check if bidNumber is a string
  if (typeof bidNumber === 'string') {
    // Extract the number from the string
    var number = parseInt(bidNumber.match(/\d+/)[0], 10);

    // If not a number set it to 0 
    if (isNaN(number)) {
      number = 0;
    }
  } else {
    number = bidNumber;
  }

  // Increment the bid number
  var newNumber = number + 1;

  // If backupSpreadsheetName is empty, set it to a default name
  if (!backupSpreadsheetName) {
    backupSpreadsheetName = ss.getName() + " Bid " + number;
  }

  // Show a toast notification
  ss.toast(backupSpreadsheetName, "🔨 Creating backup file...", 20)

  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !userSubscriptions.includes("PaidUsers")) {
    handleAccessDenied();
  } else {
    // User is a "PaidUsers" subscriber or the subscription checking is not enabled.
      
    // Make a copy of the spreadsheet with the old name + old bid number
    var backupSpreadsheetFile = Drive.Files.copy({title: backupSpreadsheetName}, ss.getId());

    // Get the project name from the named range "Project"
    var projectNameRange = ss.getRangeByName("Project");
    var projectName = projectNameRange.getValue();

    // Get the OpenBID folder
    var openBIDFolder = getOpenBIDFolder();

    // Create the folders  
    var projectFolder = findOrCreateFolder(openBIDFolder, projectName);
    var bidBackupFolder = findOrCreateFolder(projectFolder, 'OpenBID Backups');

    // 'Move' the backup spreadsheet to the target folder
    Drive.Files.update(
      { 'parents': [ { 'id': bidBackupFolder.getId() } ] }, 
      backupSpreadsheetFile.id
    );

    // Update the client summary sheet and bid number
    updateClientSummary(ss, newNumber);

    // Show a toast notification
    ss.toast("Made a backup file, updated the client summary and versioned up the bid.","🎉 Done!", 10);
  }
}


/**
 * This Apps Script function exports specified sheets from the active Google Spreadsheet 
 * as separate CSV files. It places these files into a specific directory structure
 * in Google Drive. The directories are created if they don't already exist.
 */
function downloadSheetsAsCSV() {
    // Log user activity
  logUserActivity("Expoting CSV files");

  // Get the current active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Extract the name of the spreadsheet for use in naming the CSV files
  var spreadsheetName = ss.getName(); 

  // An array of the names of the sheets that we want to export as CSV files
  var sheets = ['EXPORT ASSETS', 'EXPORT SHOTS']; 

  // Check if the sheets exist
  for (var i = 0; i < sheets.length; i++) {
    if(ss.getSheetByName(sheets[i]) == null) {
      // Show a toast message and exit
      ss.toast("You can only export from a standard OpenBID spreadsheet.", "⚠️ Error");
      return;
    }
  }

  // Check the subscription status of the user
  var userSubscriptions = isUserSubscribed();

  // If subscription checking is enabled and the user is not a "PaidSups" or "PaidUsers" subscriber, deny access.
  if (ENABLE_SUBSCRIPTION_CHECK && !(userSubscriptions.includes("PaidSups") || userSubscriptions.includes("PaidUsers"))) {
    handleAccessDenied();
  } else {
    // User is a "PaidSups" or "PaidUsers" subscriber or the subscription checking is not enabled.

    // Get the project name from the named range "Project"
    var projectNameRange = ss.getRangeByName("Project");
    var projectName = projectNameRange.getValue();

    // Get the OpenBID folder
    var openBIDFolder = getOpenBIDFolder();

    // Create the folders  
    var projectFolder = findOrCreateFolder(openBIDFolder, projectName);
    var folder = findOrCreateFolder(projectFolder, "CSV");

    // Iterate over each sheet specified in the 'sheets' array
    for (var i = 0; i < sheets.length; i++) {
      var sheet = ss.getSheetByName(sheets[i]); // Get the sheet object

      ss.toast("📤 Creating " + spreadsheetName + " " + sheets[i] + ".csv","Working..");
        
      // Create a temporary Spreadsheet because SpreadsheetApp allows only the entire spreadsheet to be exported to CSV, not individual sheets
      var tempSpreadsheet = SpreadsheetApp.create("Temp Spreadsheet", sheet.getMaxRows(), sheet.getMaxColumns());
      var tempSheet = tempSpreadsheet.getSheets()[0];

      // Get the range covering all cells in the source sheet
      var range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
              
      // Copy the values from the source sheet to the temporary spreadsheet
      tempSheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setValues(range.getValues());
              
      // Get the temporary Spreadsheet's ID
      var tempSpreadsheetId = tempSpreadsheet.getId();

      // Use UrlFetchApp to get the CSV data of the temporary Spreadsheet
      var url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheetId + "/export?format=csv";
      var params = {method:"GET", headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
      var csvData = UrlFetchApp.fetch(url, params).getBlob();

      // Create a new file in the folder with the CSV data
      folder.createFile(csvData.setName(spreadsheetName + " " + sheets[i] + '.csv'));
              
      // Delete the temporary Spreadsheet
      DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    }
    ss.toast("See " + ss.getName() + " " + OPENBID_ROOT_FOLDER + " -> " + projectName + " -> CSV ", "📂 Done! ", 30);
  }
}


/**
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 *
 * Helper functions
 * 
 * --------------------------------------------------------------------------------------------------------------------------------------------------------
 */


/**
 * This function generates a hash code from a string input.
 * It's a simple, fast hash function that should be suitable for cache key generation.
 * It's not intended for cryptographic or security purposes.
 * It converts each character of the string to a 32-bit integer, 
 * and then uses bitwise shifting and addition operations to generate the hash.
 *
 * @param {string} s - The string to be hashed.
 * @return {string} - The hash code of the input string.
 */
function hash(s) {
  // Initialize the hash to 0
  var hash = 0;

  // If the input string is empty, return the hash
  if (s.length == 0) return hash;

  // Iterate over each character in the string
  for (i = 0; i < s.length; i++) {
    // Get the character code
    var char = s.charCodeAt(i);

    // Update the hash using bitwise shifting and addition
    hash = ((hash<<5)-hash)+char;

    // Convert to a 32-bit integer
    hash = hash & hash;
  }

  // Convert the hash to a string and return it
  return hash.toString();
}


/**
 * Helper function to generate a PDF from a specified Google Sheet.
 * The PDF will have the following settings:
 * - Exported as PDF
 * - Letter paper size
 * - Portrait orientation
 * - Fit to width
 * - Sheet names, print titles, and page numbers hidden
 * - Gridlines hidden
 *
 * @param {GoogleAppsScript.Spreadsheet} - The spreadsheet we're in.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to be converted to a PDF.
 * @param {string} pdfFileName - The desired name for the generated PDF file.
 * @returns {GoogleAppsScript.Drive.File} - The created PDF file.
 */
function generatePdfFromSheet(ss, sheet, pdfFileName) {
  // Construct the export URL for the specified sheet with the desired settings
  var url = 'https://docs.google.com/spreadsheets/d/'
    + ss.getId() + '/export?'
    + 'format=pdf' // export as pdf
    + '&size=letter' // paper size
    + '&portrait=true' // orientation, false for landscape
    + '&fitw=true' // fit to width, false for actual size
    + '&sheetnames=false&printtitle=false&pagenumbers=false' // hide optional UI elements
    + '&gridlines=false' // hide gridlines
    + '&gid=' + sheet.getSheetId(); // the sheet's Id

  // Get the OAuth token for the current user
  var token = ScriptApp.getOAuthToken();

  // Fetch the PDF using the export URL and the user's OAuth token
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  // Create a blob from the fetched PDF data and set its name
  var pdfBlob = response.getBlob().setName(pdfFileName);

  // Create and return a file from the PDF blob using Drive API
  var file = Drive.Files.insert({
    title: pdfFileName,
    mimeType: 'application/pdf'
  }, pdfBlob);
  
  // Return the created file as a Google Apps Script Drive File object
  return DriveApp.getFileById(file.id);
}


/**
 * Helper function to generate a multipage PDF from an entire Google Sheets Spreadsheet.
 * The PDF will have the following settings:
 * - Exported as PDF
 * - Letter paper size
 * - Portrait orientation
 * - Fit to width
 * - Sheet names, print titles, and page numbers hidden
 * - Gridlines hidden
 *
 * @param {GoogleAppsScript.Spreadsheet} - The spreadsheet we're in.
 * @param {string} pdfFileName - The desired name for the generated PDF file.
 * @returns {GoogleAppsScript.Drive.File} - The created PDF file.
 */
function generatePdfFromSpreadsheet(ss, pdfFileName) {
  try {
    // Log the start of PDF generation
    Logger.log("Starting PDF generation...");

    // Construct the export URL for the entire spreadsheet with the desired settings
    var url = 'https://docs.google.com/spreadsheets/d/'
      + ss.getId() + '/export?'
      + 'format=pdf' // export as pdf
      + '&size=letter' // paper size
      + '&portrait=false' // orientation, false for landscape
      + '&fitw=true' // fit to width, false for actual size
      + '&sheetnames=false&printtitle=false&pagenumbers=false' // hide optional UI elements
      + '&gridlines=false'; // hide gridlines

    // Get the OAuth token for the current user
    var token = ScriptApp.getOAuthToken();

    // Fetch the PDF using the export URL and the user's OAuth token
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true // Add this line to see the full error response
    });

    // If the response code is not 200 (OK), throw an error
    if (response.getResponseCode() !== 200) {
      ss.toast("Failed to fetch PDF");
      throw new Error('Failed to fetch PDF: response code ' + response.getResponseCode());
    }

    // Create a blob from the fetched PDF data and set its name
    var pdfBlob = response.getBlob().setName(pdfFileName);

    // Create and return a file from the PDF blob using Drive API
    var file = Drive.Files.insert({
      title: pdfFileName,
      mimeType: 'application/pdf'
    }, pdfBlob);
  
    // Log successful completion
    ss.toast("📄 PDF created!");
    Logger.log("PDF generation completed successfully!");

    // Return the created file as a Google Apps Script Drive File object
    return DriveApp.getFileById(file.id);

  } catch (error) {
    // Log any errors
    ss.toast("⚠️ Error generating PDF!");
    Logger.log('Error generating PDF: ' + error.message);
    throw error;  // Re-throw the error to be handled by the calling function
  }
}


/**
 * This function checks if a document property already exists for 'openBIDFolderId'.
 * If it does, the corresponding folder is retrieved. If the folder no longer exists (perhaps it was deleted),
 * a new folder is created and its ID is stored as a document property for future use.
 * If the document property doesn't exist, it also creates a new folder and stores its ID.
 *
 * @return {Folder} - The existing or newly created OpenBID folder.
 */
function getOpenBIDFolder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if a document property already exists for 'openBIDFolderId'
  var documentProperties = PropertiesService.getDocumentProperties();
  var openBIDFolderId = documentProperties.getProperty('openBIDFolderId');

  var openBIDFolder;
  
  if (openBIDFolderId) {
    // If the document property exists, try to retrieve the corresponding folder
    try {
      openBIDFolder = DriveApp.getFolderById(openBIDFolderId);
    } catch(e) {
      // The folder has been deleted or the ID is invalid. Reset openBIDFolderId
      openBIDFolderId = null;
    }
  }

  if (!openBIDFolderId) {
    // If the document property doesn't exist or the folder has been deleted, create the necessary folders in the root directory.
    openBIDFolder = findOrCreateFolder(null, ss.getName() + ' ' + OPENBID_ROOT_FOLDER);
    
    // Save the ID of openBIDFolder to the document properties
    documentProperties.setProperty('openBIDFolderId', openBIDFolder.getId());
  }

  return openBIDFolder;
}


/**
 * Finds a folder with the given name within the root folder / parent folder or creates one if it doesn't exist.
 *
 * @param {Folder} parentFolder - The parent folder in which to search for or create the folder.
 * @param {string} folderName - The name of the folder to search for or create.
 * @return {Folder} - The existing or newly created folder with the given name.
 */
function findOrCreateFolder(parentFolder, folderName) {
  // If no parentFolder is provided, default to the root of Google Drive
  var parentFolderId = parentFolder ? parentFolder.getId() : 'root';

  // Find folders with the given name in the parent folder using Drive API
  var response = Drive.Files.list({
    q: "mimeType='application/vnd.google-apps.folder' and trashed=false and '" + parentFolderId + "' in parents and title='" + folderName + "'",
    fields: 'items(id)',
  });

  // If a folder with the given name already exists, return it
  if (response.items.length > 0) {
    return DriveApp.getFolderById(response.items[0].id);
  } else {
    // If a folder with the given name doesn't exist, create one and return it
    var newFolder = Drive.Files.insert({
      title: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [{id: parentFolderId}]
    });
    return DriveApp.getFolderById(newFolder.id);
  }
}



/**
 * Returns the last row in a specified column that contains data.
 *
 * @param {Sheet} sheet - The sheet object.
 * @param {number} column - The column number (1-indexed).
 * @return {number} The last row with data in the specified column.
 */
function getLastRowWithData(sheet, column) {
  var data = sheet.getDataRange().getValues();
  var lastRow = data.length;
 
  for (var i = lastRow - 1; i >= 0; i--) {
    if (data[i][column - 1]) {
      return i + 1;
    }
  }
 
  // Return the result object.
  return 0;
}


/**
 * Logs user activity in the LOGS sheet of the specified spreadsheet.
 * The logged information includes the current date and time, userEmail, and action.
 *
 * @param {string} action - A description of the action performed by the user.
 */
function logUserActivity(action) {
  // Get the user
  var userEmail = Session.getEffectiveUser().getEmail();
 
  // Get the LOGS sheet from the specified spreadsheet
  var logsSheet = SpreadsheetApp.openById(USERLOG_SHEET_ID).getSheetByName("LOGS");

  // Check if the "LOGS" sheet exists in the spreadsheet
  if (!logsSheet) {
    // Create a new "LOGS" sheet if it doesn't exist
    logsSheet = SpreadsheetApp.openById(USERLOG_SHEET_ID).insertSheet("LOGS");
  }

  // Get the current date and time
  var currentDate = new Date();

  // Append the currentDate (including time), userEmail, and action to the "LOGS" sheet
  logsSheet.appendRow([currentDate, userEmail, action]);
}


/**
 * Function to open the get vendor Google Form in a new browser tab.
 * This function displays a custom dialog with the form within the Google Sheets UI.
 */
function openGetVendorFormInPopup() {
  // Log user activity
  logUserActivity("Get vendor form shown");

  // The iframe HTML code for the Google Form
  var formIframe = '<iframe src="https://docs.google.com/forms/d/e/1FAIpQLSeJaYRoqV4NHuy99y_S6gxdSiIi_QT3aAv4y7lyao7H_R51Hw/viewform?embedded=true" width="640" height="2450" frameborder="0" marginheight="0" marginwidth="0">Loading…</iframe>';
 
  // Create an HTML output with the Google Form iframe
  var htmlOutput = HtmlService.createHtmlOutput(formIframe)
    .setWidth(670) // Set the width of the popup window
    .setHeight(650); // Set the height of the popup window

  // Display the custom dialog with the Google Form within the Google Sheets UI
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}


/**
 * Function to open the vendor application Google Form in a new browser tab.
 * This function displays a custom dialog with the form within the Google Sheets UI.
 */
function openVendorApplicationFormInPopup() {
  // Log user activity
  logUserActivity("Become a vendor form shown");

  // The iframe HTML code for the Google Form
  var formIframe = '<iframe src="https://docs.google.com/forms/d/e/1FAIpQLScqqQ9f7U8Z8vmOKgWCp7Ab-UIAVBVun7AqTMcdMCRdRMe0vw/viewform?embedded=true" width="640" height="660" frameborder="0" marginheight="0" marginwidth="0">Loading…</iframe>';
 
  // Create an HTML output with the Google Form iframe
  var htmlOutput = HtmlService.createHtmlOutput(formIframe)
    .setWidth(670) // Set the width of the popup window
    .setHeight(650); // Set the height of the popup window

  // Display the custom dialog with the Google Form within the Google Sheets UI
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}


/**
 * Function to open the voting Google Form in a new browser tab.
 * This function displays a custom dialog with the form within the Google Sheets UI.
 */
function openFormInPopup() {
  // The iframe HTML code for the Google Form
  var formIframe = '<iframe src="https://docs.google.com/forms/d/e/1FAIpQLSf936fHpIjng-99MYIMHnHwh5JUQ1OtqsgIBYO2QMSEtmp6mg/viewform?embedded=true" width="600" height="4000" frameborder="0" marginheight="0" marginwidth="0">Loading…</iframe>';
 
  // Create an HTML output with the Google Form iframe
  var htmlOutput = HtmlService.createHtmlOutput(formIframe)
    .setWidth(660) // Set the width of the popup window
    .setHeight(650); // Set the height of the popup window

  // Display the custom dialog with the Google Form within the Google Sheets UI
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}


/**
 * A custom function to determine the last row in a Google Sheet column that contains data.
 * This function addresses the limitation of getLastRow() which returns the last row that contains 
 * any data, even if there are empty rows in between. This function ensures we get the last row 
 * specifically in the desired column.
 *
 * @param {Sheet} sheet - A Google Apps Script Sheet object representing the sheet where the operation is to be performed.
 * @param {string} column - A string denoting the column (e.g., "A", "B") where we need to find the last row with data.
 * @returns {number} - The row number of the last row with content in the specific column. If no content is found in the column, it returns -1.
 */
function getLastRowWithContent(sheet, column) {
  // Logger.log("Retrieving the last row with content in column: " + column);

  // Create a range that spans the entire column. The range is represented as "A:A", "B:B", etc.
  var range = sheet.getRange(column + ":" + column);
  // Logger.log("Created range for column: " + column);

  // Retrieve all values from the defined range.
  // As our range is a single column, getValues() will return a 2-dimensional array where each inner array has a single value, corresponding to a cell in that column.
  var values = range.getValues();
  // Logger.log("Retrieved values from column: " + column);

  // Initialize the lastRow variable to -1. This will be the default return value if no content is found in the column.
  var lastRow = -1;

  // Iterate from the bottom of the column (values.length - 1) to the top (0) to find the last row with content.
  for (var i = values.length - 1; i >= 0; i--) {
    // If the current cell (values[i][0]) is not empty
    if (values[i][0] != "") {
      // Update lastRow to the current row number. We add 1 to 'i' because row numbers in Sheets start from 1 whereas array indexes in JavaScript start from 0.
      lastRow = i + 1;
      // Logger.log("Found last row with content at row: " + lastRow);

      // As we have found the last row with content, we break the loop and no further checking is required.
      break;
    }
  }
  
  // Return the row number of the last row with content. If no content was found, it will return the initial value of -1.
  // Logger.log("Returning last row with content: " + lastRow);
  return lastRow;
}


/**
 * Adjusts an array of formulas to match a target row by replacing the row references.
 *
 * @param {Array} formulas - The array of formulas to adjust.
 * @param {number} targetRow - The target row number.
 * @returns {Array} - The adjusted formulas.
 */
function adjustFormulasRow(formulas, targetRow) {
  var adjustedFormulas = []; // The array to store the adjusted formulas
  
  // Loop through each formula in the input array
  for (var i = 0; i < formulas.length; i++) {
    // If the formula is not an empty string
    if (formulas[i] !== "") {
      // Use a regular expression to replace cell references in the formula with the target row number
      // The regular expression /(\$?[A-Z]+\d+)/g matches cell references which optionally start with $ (e.g., $A1, B1, etc.)
      var adjustedFormula = formulas[i].replace(/(\$?[A-Z]+\d+)/g, function(match, p1) {
        // The prefix is the $ if the matched cell reference starts with $
        var prefix = p1.startsWith("$") ? "$" : "";
        
        // The cell reference without the prefix (if any)
        var ref = p1.startsWith("$") ? p1.substring(1) : p1;
        
        // The column part of the cell reference (e.g., "A" in "A1")
        var column = ref.match(/[A-Z]+/)[0];
        
        // Return the adjusted cell reference with the target row number
        return prefix + column + targetRow;
      });
      
      // Push the adjusted formula into the array
      adjustedFormulas.push(adjustedFormula);
      
      // Log the adjusted formula for debugging
      // Logger.log("Adjusted formula for index " + i + ": " + adjustedFormula);
    } else {
      // If the formula is an empty string, push an empty string into the array
      adjustedFormulas.push("");
    }
  }
  
  // Return the array of adjusted formulas
  return adjustedFormulas;
}


/**
 * This function will error out if in the OpenBID client or OpenBID Vendor Template spreadsheet by using the named ranges "firewallFile" or "VendorNumber" respectively.
 */
function checkInOpenBidVendor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get all named ranges in the spreadsheet
  var namedRanges = ss.getNamedRanges();
  var isAllCompanyInfoPresent = false;

  // Loop through all named ranges
  for (var i = 0; i < namedRanges.length; i++) {
    // Check if the named range is "firewallFile" or "VendorNumber"
    if (namedRanges[i].getName() === "firewallFile" || namedRanges[i].getName() === "VendorNumber") {
      // Throw a toast notification with the error message
      ss.toast("This tool is exclusively for standalone OpenBID Vendor spreadsheets and will not function when linked to an OpenBID Client.", "Error ⚠️", 10);
      // Return false, indicating that a disallowed named range was found
      return false;
    }
    if (namedRanges[i].getName() === "AllCompanyInfo") {
      isAllCompanyInfoPresent = true;
    }
  }
  
  if (!isAllCompanyInfoPresent) {
    ss.toast("This tool is exclusively for OpenBID spreadsheets and will not function without one.", "Error ⚠️", 10);
    return false;
  }

  // If no named ranges named "firewallFile" or "VendorNumber" are found, return true
  return true;
}



/**
 * This function hides empty rows in a specified sheet starting from the 6th row. 
 * It considers a row to be empty if the cell in column H is empty. 
 * The hiding operation is performed from the bottom to the top.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet in which to hide empty rows.
 */
function hideEmptyRows(sheet) {
  // Log the start of the row hiding operation
  Logger.log("Starting row hiding operation...");

  // Get all the rows in the sheet.
  var rows = sheet.getDataRange().getValues();
  Logger.log("Fetched all rows from the sheet.");

  // Determine the number of rows in the sheet.
  var numRows = rows.length;
  Logger.log("Determined number of rows in the sheet: " + numRows);

  // Iterate through the rows in reverse order, starting from the last row and ending at the 5th row.
  for (var i = numRows - 1; i >= 5; i--) {
    // Get the data in the current row.
    var row = rows[i];

    // Get cell value in column H (index 7)
    var cellValueH = row[7];

    // If the cell is empty, hide the row.
    if (!cellValueH) {
      sheet.hideRows(i + 1);  // The hideRows function is 1-indexed, so we add 1 to the loop counter.
      Logger.log("Hid row: " + (i + 1));
    }
  }

  // Log the completion of the row hiding operation
  Logger.log("Completed row hiding operation.");
}


/**
 * This function updates the Client Summary in a given spreadsheet.
 * It retrieves the values of the Client Summary 
 * and pastes the values into row 11.
 * Finally, it increments the bid number in the 'BidNumber' named range.
 *
 * @param {Spreadsheet} ss - The spreadsheet to update
 * @param {number} newNumber - The new bid number
 */
function updateClientSummary(ss, newNumber) {
  // Retrieve the "ClientSummary" named range
  var clientSummaryRange = ss.getRangeByName("ClientSummary");
  if (!clientSummaryRange) {
    ss.toast("Client Summary not found", "⚠️ Error");
    throw new Error('Named range "ClientSummary" not found');
  }

  // Retrieve the values of the "ClientSummary" range
  var clientSummaryValues = clientSummaryRange.getValues();

  // Get the "SUMMARY" sheet
  var summarySheet = ss.getSheetByName('SUMMARY');

  // Insert a new row below row 10
  summarySheet.insertRowsAfter(10, 1);

  // Get the range in the new row where to paste the "ClientSummary" values and formatting
  var newRange = summarySheet.getRange(11, clientSummaryRange.getColumn(), 1, clientSummaryRange.getWidth());

  // Copy the ClientSummary range to the new range
  clientSummaryRange.copyTo(newRange);

  // Overwrite the values in the new range with the original values from the ClientSummary range
  newRange.setValues(clientSummaryValues);

  // Retrieve the "BidNumber" named range
  var namedRange = ss.getRangeByName("BidNumber");
  if (!namedRange) {
    ss.toast("Named range BidNumber not found", "⚠️ Error");
    throw new Error('Named range "BidNumber" not found');
  }

  // Replace the value in the "BidNumber" named range with the incremented bid number
  namedRange.setValue(newNumber);
}


/**
 * Displays a popup to inform the user about the created backup spreadsheet.
 * The popup contains a red button that links to the backup spreadsheet URL.
 *
 * @param {string} backupSpreadsheetURL - URL of the created backup spreadsheet.
 * @param {number} newNumber - The updated bid version number.
 */
function exportPDFPopup(backupSpreadsheetURL, newNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Create a popup with a red button linked to the backupSpreadsheetURL using the HTML file
  var htmlTemplate = HtmlService.createTemplateFromFile('exportPDFPopup');
  htmlTemplate.backupSpreadsheetURL = backupSpreadsheetURL;
  htmlTemplate.newNumber = newNumber;
  var htmlOutput = htmlTemplate.evaluate().setWidth(450).setHeight(800);

  // Show the popup
  ui.showModalDialog(htmlOutput, 'Backup Spreadsheet Created!');
}
