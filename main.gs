// *****************
// Show UI on Open
function onOpen() {
	SpreadsheetApp.getUi()
		.createMenu('FSD Tools')
		.addItem('Create FSD Tools Folder on Drive', 'dzn_createFolderFromMenu')
		.addSeparator()
		.addItem('⌕ Parse Mission.sqm', 'dzn_mp_parseFromMenu')
		.addItem('✔ Confirm Data', 'dzn_mp_confirmParsingFromMenu')
		.addSeparator()
		.addItem('Create Sloting form COOP', 'dzn_createSlottingCoopFromMenu')
		.addItem('Create Sloting form TVT', 'dzn_createSlottingTvtFromMenu')
		.addSeparator()
		.addItem('Show Sidebar', 'showSidebar')
		.addToUi();
 }

// *****************
// Show sidebar
function showSidebar(link1, link2) {
	var htmlOut = "<style>p, li {font-size: 6;}div {background: #2E2E2E; width: 250px; border-radius: 12px; color: #FFCC00; font-family: 'Trebuchet MS', monospace; padding-left: 16px;}a {  padding-left: 30px; font-family: Consolas, monospace; border-radius: 52px; font-size: 14; text-decoration: none;}</style><p>FSD Slotting Tool - is a way to create slotting and feedback forms for your multiplayer online game.<p>";
  	var ss = SpreadsheetApp.getActive();
	if (link1 != null) {
		htmlOut = htmlOut + "<div>Newly Created Forms</div><a href='" 
			+ link1 
			+ "' target='_blank' title='"
            + ss.getRangeByName('slotName').getValue()
            + "'>Slotting Form</a>";
		if (link2 != null) {
			htmlOut = htmlOut + "<br><a href='" 
				+ link2 
				+ "' target='_blank' title='"
                + ss.getRangeByName('feedName').getValue()
                + "'>Feedback Form</a>"; 
		}
	} else {
		if (SpreadsheetApp.getActive().getRangeByName('slotURL').getValue() != "") {
			htmlOut = htmlOut + "<div>Last Created Forms</div><a href='" 
				+ ss.getRangeByName('slotURL').getValue() 
				+ "' target='_blank' title='"
                + ss.getRangeByName('slotName').getValue()
                + "'>Slotting Form</a>";
			if (SpreadsheetApp.getActive().getRangeByName('feedURL').getValue() != "") {
				htmlOut = htmlOut + "<br><a href='" 
					+ ss.getRangeByName('feedURL').getValue() 
					+ "' target='_blank' title='"
                    + ss.getRangeByName('feedName').getValue()
                    + "'>Feedback Form</a>";
			}
		}
	}

	htmlOut = htmlOut + dzn_htmlInstruction();
  
	var html = HtmlService.createHtmlOutput(htmlOut)    
		.setSandboxMode(HtmlService.SandboxMode.IFRAME)
		.setTitle('FSD Slotting Tools')
		.setWidth(300)
	SpreadsheetApp.getUi().showSidebar(html); 
}


// *****************
// Creating base folder
function dzn_createFolderFromMenu() {    
	if (DriveApp.getFoldersByName("ARMA FSD Tools").hasNext()) {
		SpreadsheetApp.getUi().alert('✔ OK\n\nFolder already exists');
	} else {        
		SpreadsheetApp.getUi().alert('Folder with name "ARMA FSD Tools" will be created in the root of your Google Drive');
		DriveApp.createFolder("ARMA FSD Tools");
		SpreadsheetApp.getUi().alert('✔ OK\n\nFolder "ARMA FSD Tools" was created successfully');
	}  
}

// *****************
// Check the existance of work folder
function dzn_checkFolderExists() {
	var output = true;
	if (!(DriveApp.getFoldersByName("ARMA FSD Tools").hasNext())) {
		SpreadsheetApp.getUi().alert('⊗ WARNING!\n\nThere is no "ARMA FSD Tools" folder on your Drive.\n\nPlease, create it via "FSD Tools" menu');    
		output = false;
	}
	return output
}

// *****************
// Creating COOP or TVT form from menu
function dzn_createSlottingCoopFromMenu() {
	if (dzn_checkFolderExists()) {
		SpreadsheetApp.getUi().alert('Starting to creating COOP Forms.\n\nPress OK and wait for a while.'); 
		dzn_createForm('COOP'); 
	}
}

function dzn_createSlottingTvtFromMenu() {
	if (dzn_checkFolderExists()) {
		SpreadsheetApp.getUi().alert('Starting to creating TVT Forms.\n\nPress OK and wait for a while.'); 
		dzn_createForm('TVT'); 
	}
}

// *****************
// Creating Slotting form for COOP
function dzn_createForm(mode) {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
  
	// Запрашиваем создание формы слоттинга
	var formSlotName = dzn_getSlotFormName(mode);   
	if (formSlotName[1] != "null") {
		// Clear URLs
		ss.getRangeByName("slotURL").clearContent();
		ss.getRangeByName("slotURL").clearContent();
      	ss.getRangeByName("feedURL").clearContent();
		ss.getRangeByName("feedName").clearContent();
      
		var slotFormUrl, feedFormUrl
		var folder =  DriveApp.getFoldersByName("ARMA FSD Tools").next().createFolder(formSlotName[0]);	 
		var formSlotId = FormApp.create(formSlotName[0]).getId();		
		folder.addFile(DriveApp.getFileById(formSlotId));
		DriveApp.getRootFolder().removeFile(DriveApp.getFileById(formSlotId)); 
		slotFormUrl = DriveApp.getFileById(formSlotId).getUrl();
      
		var propSheetId = SpreadsheetApp.create('properties ' + formSlotName[0]).getId();
		folder.addFile(DriveApp.getFileById(propSheetId));
		DriveApp.getRootFolder().removeFile(DriveApp.getFileById(propSheetId));
      
		var propSheet = SpreadsheetApp.openById(propSheetId);
      
		// Запрашиваем создание формы фидбека
		var formFeedName = dzn_isFeedbackNeeded(mode, formSlotName[1]);
		if (formFeedName != "null") {
			var formFeedId = FormApp.create(formFeedName).getId();
			folder.addFile(DriveApp.getFileById(formFeedId));
			DriveApp.getRootFolder().removeFile(DriveApp.getFileById(formFeedId));
			feedFormUrl = DriveApp.getFileById(formFeedId).getUrl();          
          
			dzn_addNamedRanges("Feed", propSheetId, ss.getId(), mode); 
			// PreInitialize Feedback Form
			dzn_feedForm_preInitialize(formFeedId, propSheet.getId()); 
        }

		dzn_addNamedRanges("Slot", propSheetId, ss.getId(), mode);
		// Pre-initialize Slotting Form
		dzn_slotForm_preInitialize(formSlotId, propSheet.getId());

		ss.getRangeByName("slotURL").setValue(slotFormUrl);
		ss.getRangeByName("slotName").setValue(formSlotName[0]);
		if (feedFormUrl != null) {			
			ss.getRangeByName("feedURL").setValue(feedFormUrl);
			ss.getRangeByName("feedName").setValue(formFeedName);
            showSidebar(slotFormUrl, feedFormUrl);
		} else {
			showSidebar(slotFormUrl)
		}

		SpreadsheetApp.getUi().alert('✔ OK\n\nForm was successufully created!\n\nCheck Sidebar for URLs');   
	};
}


// *****************
// Create PROMPT and return the answer for SLOTTING FORM
function dzn_getSlotFormName(gametype) {
	var ui = SpreadsheetApp.getUi();
	var result = "";
	var output = ["null"];
	
	result = ui.prompt('Create Form - Mission Title','Please enter the Mission Title:', ui.ButtonSet.OK_CANCEL);
	var button = result.getSelectedButton();
	var text = result.getResponseText();
	if (button == ui.Button.OK) {
		if (text.length == 0) {
			text = dzn_getToday();
		}
		output = [(gametype + ' ' + text), text];
	}
	return output
}

// *****************
// Create PROMPT and return YES/NO to create FEEDBACK
function dzn_isFeedbackNeeded(gameType, gameName) {
	var ui = SpreadsheetApp.getUi();	
	var output = "null";
	var result = ui.alert('Do you want to create Feedback form?', ui.ButtonSet.YES_NO);
	if (result == ui.Button.YES) {
		output = gameType + ' AAR ' + gameName;
	}
	return output
}

// *****************
// Return today date
function dzn_getToday() {
	var today = new Date();
	var dd = today.getDate();
	var mm = today.getMonth()+1; //January is 0!
	var yyyy = today.getFullYear();
	if(dd<10) {dd='0'+dd};
	if(mm<10) {mm='0'+mm};
	today = mm+'/'+dd+'/'+yyyy;
	return today
}
