/*

This script is used in conjunction with a lending library spreadsheet. It runs everytime a new submission from an attached from comes in. The form must
First be configured withe right questions/columns. Each time a 'client' checks out an item that is scanned, the script checks if the form was a return or
checkout. In case of a checkout, it updates the items database with new last user, checkout data, location, and due date.



//USAGE
 
  
  1.Create a new Google Sheet,
  2. Create Form from sheet
  3. Go to Tools > Script editor and launch it. Approve access. Set Project triggers to onFormSubmit
  
  
  */
 
 ///Constants....Please update to match your sheet. Remember, the columns start at 0 with the timestamp.The second column is indexed as 1...
 
 //FROM FORM SUBMISSION TAB...
 
 var EMAIL_COLUMN = 4;
 var FULL_NAME_COLUMN = 1;
 var SCHOOL_COLUMN = 2;
 var SCHOOL_ADDRESS = 3;
 var PHONE_COLUMN = 5;
 var USER_ROLE_COLUMN_ON_FORM_SUBMIT = 6;
 var CHECKING_OR_RETURNING_COLUMN = 11;
 var CHECKING_OUT_ITEMS_COLUMN = 13;
 var RETURNED_ITEMS_COLUMN = 12;
 
 //FROM USER DATABASE TAB ... Remember, Column A is indexed at zero

var USER_ID_COLUMN = 0; 
var USER_NAME_COLUMN = 1;
var USER_ROLE_COLUMN = 2;
var USER_ORG_COLUMN = 3; 
var  USER_ADDRESS_COLUMN = 4;
var USER_PHONE_COLUMN = 5;
var USER_EMAIL_COLUMN = 6;

//Below is the fusion table if plotting a map using the Fusion Tables service. 

var TABLE_ID = '1yt90oI79_JRmpgsZXqPHZD6JS8Xv_sO2nN4Bol9J';

// First row that has data, as opposed to header information
var FIRST_DATA_ROW = 4;

// True means the spreadsheet and table must have the same column count
var REQUIRE_SAME_COLUMNS = false;
 
 
 //adds the user Interface
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('LendingLib')
      .addItem('Send Overdue Notices', 'checkOverdue')
      .addToUi();
}

//this function runs on each form submission
function onFormSubmit(e) {


  var ss = SpreadsheetApp.getActiveSpreadsheet();


//make sure to set Triggers >>Edit >> Current Project Triggers >> Set trigger to run on onFormSubmit for both script and sheet. 

  Logger.log("A real form has been submited");
  
  //get active spreadhseet:

   var sheet = ss.getSheets()[0];
   var lastRow = sheet.getLastRow(); //get index of last row...
   
    var data = sheet.getDataRange().getValues();
   
   
    //Now decide if we are returning or checking out materials:
    
    var returnOrCheckingOut = sheet.getRange(lastRow, CHECKING_OR_RETURNING_COLUMN +1 ).getValues();
     Logger.log("Returning Materials? " + returnOrCheckingOut );
    
    if (returnOrCheckingOut == "Returning materials"){
    
    
       var dateNow = new Date(); //a date object for NOW, whenever that now happens to be. 
       
       var dateInMS = dateNow.getTime(); //grab the milliseconds value
       
        var returnedItems = data[lastRow-1][RETURNED_ITEMS_COLUMN];
      // var returnedItems = sheet.getRange(lastRow, RETURNED_ITEMS_COLUMN +1 ).getValues();
        
        Logger.log("Returning Materials : " + returnedItems );
        
        //run returnItems
        
        var returnTransaction = [returnedItems, dateInMS];
        
         returnItems(returnTransaction);
         
         sync();
        
        return;
    
    } 
    
    
   
   //grab the email from the fifth column...(make sure it matches)
   
   var range = sheet.getRange(lastRow, EMAIL_COLUMN +1); //this is the email column, AT THIS STAGE, ITS INDEXED STARTING AT 1...
   
   var range2 = sheet.getRange(lastRow,1, 1, 11);
   
   var values = range.getValues(); //gets the email value
   var values2 = range2.getValues(); //gets the whole row
   
   var submitterEmail = values[0][0];
  
   //send info to check user
   var thisUser = getUserInfo(submitterEmail);
   
   
   Logger.log("user ID of person checking out: " + thisUser);
   
   /********************
       
       Below makes sure columns match the second []...
       
       **********************/
       
        var fullName = data[lastRow-1][FULL_NAME_COLUMN];
        
        var schoolOrg = data[lastRow-1][SCHOOL_COLUMN];
        
        var schoolAddress = data[lastRow-1][SCHOOL_ADDRESS];
        
        var userEmail = data[lastRow-1][EMAIL_COLUMN];
        
        var userPhone = data[lastRow-1][PHONE_COLUMN];
        
        var userRole = data[lastRow-1][USER_ROLE_COLUMN_ON_FORM_SUBMIT];
        
        var date = data[lastRow-1][0]; //grab the timestamp from col. 0
        
        var items = data[lastRow-1][CHECKING_OUT_ITEMS_COLUMN];
        
 
        
        Logger.log("Items to check out are: " + items);
   
     if(thisUser == 0){
       //we have a new user, so let's create an entry
       
        
        var newUser =[
               
                fullName,
                schoolOrg,
                schoolAddress,
                userEmail,
                userPhone,
                userRole
              ];    
          
          
          thisUser = createUser(newUser);
          
          Logger.log("created new user! and their ID is: " + thisUser);
    
    } else {
    
      //a returning customer!
      Logger.log("We have a returning customer! " + thisUser);
      
     }
     
     
     //look at last row and get the information we need for to update the record, which is user, address, date and item array. 
     
       var transaction =[
               
                thisUser,
                schoolAddress,
                items,
                date,
                
              ];    
              
       Logger.log("Ready to update items...");
          
      updateItems(transaction); //send it off to be updated
      
       Logger.log("Items updated...");
       
       
     sync(); //for fusion tables integration
     
     Logger.log("Synced with Fusion Tables");

}



function getUserInfo(email){

//Logger.log(val);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("Clients");
  var data = userSheet.getDataRange().getValues();
   
  //determine if user already in the system
  //look through the column and find user by email...
  
  var numRows = userSheet.getLastRow(); // returns the index of last row
  
  var userID = 0;
  
    for (var i=1; i<numRows; i++){
      
           if(data[i][USER_EMAIL_COLUMN] == email ){
          
           //if in the system, set userID as currentID
           
            userID = data[i][0];
            
            return userID;
           
          } else {
          
       
          
          }
        
        } //end for
  
    Logger.log("No user found, setting USERID to 0");
    
    return userID;

}

function updateItems(record){


   Logger.log("Updating record : " + record);
   Logger.log("setting up sheet ...");
 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var itemSheet = ss.getSheetByName("Items");
    var data = itemSheet.getDataRange().getValues();
    
    
    var userSheet = ss.getSheetByName("Clients");
    var userData = userSheet.getDataRange().getValues();
    var numberOfClients = userSheet.getLastRow();
    
    var numRows = itemSheet.getLastRow();
    
    Logger.log("[updateItems] There are these many users: " + numberOfClients);
    
    var settingsSheet = ss.getSheetByName("Settings");
    
    var lendingPeriod = settingsSheet.getRange(4,4).getValues(); //see settings tab in spreadsheet
    
    
    var user = record[0];
    
    var shoppingCart = [];
    
    //Logger.log("User to be sent reminder: " + user );
    
    var location = record[1];
    var items = record[2]; //could be an array, if multiple items
    var date = record[3]; //might remove and instead use dateNow
    
    
     var dateNow = new Date(); //a date object for NOW, whenever that now happens to be. 
     
     var dateInMS = dateNow.getTime(); //grab the milliseconds value
     
     
     
     var dueDate = new Date(dateInMS + 86400000*lendingPeriod).toDateString(); // add the number of days chosen by user for loan period create a due date
      
     dateNow = dateNow.toDateString();  //get a human readable value to populate the cell. 
      
    
    var statusCell = itemSheet.getRange(1, 3); 
    
    var locationCell = itemSheet.getRange(1, 5); 
    var lastUserCell = itemSheet.getRange(1, 4); 
    var checkedOutDateCell = itemSheet.getRange(1, 6); 
    var dueDateCell = itemSheet.getRange(1, 7); 
    
    Logger.log("Items -----> : "+ items);
    Logger.log("Items count: " + items.length);
    
    //need to handle one item vs. many
    
    if(items.length >1){
    
        Logger.log("we have more than one item!");
        
        var itemsArr = items.split("\n");
        
          //search for each item in items...
     
        for (var i=0; i<itemsArr.length; i++){
    
            
      
         for (var j=1; j<numRows; j++){
     
              //search down the 0 column, with IDs...
          
              if(data[j][0] == itemsArr[i] ){
       
              Logger.log("We have a match!! Item# " + itemsArr[i] + " == " + data[j][0]);
              
                statusCell = itemSheet.getRange(j+1, 4); 
                statusCell.setValue("Checked Out");
                
                locationCell = itemSheet.getRange(j+1, 6);
                locationCell.setValue(location);
                
                lastUserCell = itemSheet.getRange(j+1, 5);
                lastUserCell.setValue(user);
                
                checkedOutDateCell = itemSheet.getRange(j+1, 7);
                checkedOutDateCell.setValue(dateNow);
                
               
                //
                dueDateCell = itemSheet.getRange(j+1, 8);
                dueDateCell.setValue(dueDate);
                
                //now load item name into shopping car
                
                shoppingCart.push(data[j][1]); 
              
               Logger.log("Updated Item at index " + j);
              //upate the row
           
             } 
    
         }
       }
       
       
    }else{
    
        Logger.log("we have just one item to update");
        
        
         for (var j=1; j<numRows; j++){
     
              //search down the 0 column, with IDs...
          
              if(data[j][0] == items ){
       
              Logger.log("We have a match!! Item# " + items + " == " + data[j][0]);
              
                //update status
                statusCell = itemSheet.getRange(j+1, 4); 
                statusCell.setValue("Checked Out");
                
                //set the new location
                locationCell = itemSheet.getRange(j+1, 6);
                locationCell.setValue(location);
                
                //set the last user
                lastUserCell = itemSheet.getRange(j+1, 5);
                lastUserCell.setValue(user);
                
                //update checked out date
                checkedOutDateCell = itemSheet.getRange(j+1, 7);
                checkedOutDateCell.setValue(dateNow);
                
               
                //update due date
                dueDateCell = itemSheet.getRange(j+1, 8);
                dueDateCell.setValue(dueDate);
              
               Logger.log("Updated Item at index " + j);
               
               shoppingCart.push(data[j][1]); 
             
           
             } 
             
         }
    
    }
    
  
    
    for (var n= 0 ; n < numberOfClients ; n++){
           
             if (userData[n][0]== user){
             
               Logger.log("We have identified a user, their email is : " + userData[n][6] + " and their name : " + userData[n][1] );
               
               sendConfirmationEmail(userData[n][1], userData[n][6], shoppingCart, dueDate);
             
             } else {
              Logger.log("No, the email is and thus not a match : " + userData[n][6]);
             
             }
           
           
           }
  
}

function createUser(userInfo){


  //take an array of information and add it to the "users" sheet
  
  Logger.log("Creating User  : " + userInfo[3]); //logs the address
  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //load the right sheet
  var userSheet = ss.getSheetByName("Clients");
  var data = userSheet.getDataRange().getValues(); //grab all the data
 

    var lastRowIndex = userSheet.getLastRow();  //get the number of rows there are for indexing
    
    Logger.log("lastRowIndex = " + lastRowIndex);
   
   // match cells in last row with their respective data
   var idcell = userSheet.getRange(lastRowIndex+1, USER_ID_COLUMN+1); 
   var namecell = userSheet.getRange(lastRowIndex+1, USER_NAME_COLUMN+1);
   var rolecell = userSheet.getRange(lastRowIndex+1, USER_ROLE_COLUMN+1);
   var schoolcell = userSheet.getRange(lastRowIndex+1, USER_ORG_COLUMN+1);
   var addresscell = userSheet.getRange(lastRowIndex+1, USER_ADDRESS_COLUMN+1);
   var emailcell = userSheet.getRange(lastRowIndex+1, USER_EMAIL_COLUMN+1);
   var phonecell = userSheet.getRange(lastRowIndex+1, USER_PHONE_COLUMN+1);
   
   

     
     //set the values from the info passed 
     idcell.setValue(lastRowIndex);
     namecell.setValue(userInfo[0]);
     rolecell.setValue(userInfo[5]);
     schoolcell.setValue(userInfo[1]);
     addresscell.setValue(userInfo[2]);
     emailcell.setValue(userInfo[3]);
     phonecell.setValue(userInfo[4]);
     
 //send email notification of created user
 
    sendEmailToNewUser(userInfo[0],userInfo[3]);
    
    
    
    Logger.log("Success!Created User  : " + userInfo[3]); //logs the address
    
    
    //now delete row if duplicate
    
    
       var previousEmailCell = userSheet.getRange(lastRowIndex, USER_EMAIL_COLUMN+1).getValue();
     
     Logger.log("the previous email cell conntains: " + previousEmailCell );
     
     if(previousEmailCell == userInfo[3]){
     
         Logger.log("we have a dupe record");
         
         var recordToDeleteRange = userSheet.getRange(lastRowIndex+1,1, lastRowIndex+1, USER_EMAIL_COLUMN+1 );
         recordToDeleteRange.clearContent();
       return lastRowIndex-1;
     
     } else{
       Logger.log("No, we don't have a dupe record");
       return lastRowIndex;
     
     }
    
     

}


function sendEmailToNewUser(name, email){
      
     var ss = SpreadsheetApp.getActiveSpreadsheet(); //load the right sheet
     var settingsSheet = ss.getSheetByName("Settings");
     var checkboxcontent = settingsSheet.getRange(9,3).getValues();
     
     Logger.log("Checkbox for email to new user ? : " + checkboxcontent);
     Logger.log(typeof checkboxcontent);
     
     //check settings for checkbox TRUE to see if welcome email is on
     
     if(checkboxcontent == "true" ){

        var emailSubjectToNewUser = settingsSheet.getRange(9,4).getValues();
        var emailBodyToNewUser = settingsSheet.getRange(9,5).getValues();
        
        var replyToEmail = settingsSheet.getRange(9,9).getValues();
        
        Logger.log("send email to: " + name + " " + email);
        
         MailApp.sendEmail(email,replyToEmail,
                   emailSubjectToNewUser,
                   "Hello " + name + ", " + emailBodyToNewUser);
                   
         Logger.log("email sent!");

    } else{

        Logger.log("Not sending welcome email..");

    }
  
    
}

function sendConfirmationEmail(name, email, cart, dueDate){

     var ss = SpreadsheetApp.getActiveSpreadsheet(); //load the right sheet
     var settingsSheet = ss.getSheetByName("Settings");
     var checkboxcontent = settingsSheet.getRange(9,2).getValues();
     var itemsSheet = ss.getSheetByName("Items");
     var itemsData = itemsSheet.getDataRange().getValues();
     
     var itemsLastRow = itemsSheet.getLastRow();
     
     Logger.log("ItemsLastRow = " + itemsLastRow);
     
      
     Logger.log("shoppingCArt has = " + cart + " and has length of " + cart.length);
     
     
     var subject = settingsSheet.getRange(13,4).getValues();
     
     var replyToEmail = settingsSheet.getRange(9,9).getValues();
     var emailBody = settingsSheet.getRange(13,5).getValues();
     
     
     var cartExpandedString = '';
     
     for(var i = 0; i<cart.length ; i++) {
     
         cartExpandedString = cartExpandedString +     cart[i] + ", "; 
       
     }
     
      
     emailBody =  " \n" + emailBody + " \n \n ITEMS : \n " + cartExpandedString + " \n \n due on : \n" + dueDate ;
     //only send if specified...
     
     
    Logger.log("sendConfEmail...w/body :" + emailBody);
     
      if(checkboxcontent == "true"){
      
      //compose email

        Logger.log("Sending email confirmation to: " + name + " at : " + email + " items : " + cart + " which are due on " + dueDate);
         
         MailApp.sendEmail(email,replyToEmail,
                   subject,
                   "Hello " + name + ", " + emailBody);
                   
                   
                   
         Logger.log("email sent!");

      //send mail
      
      
      }
     

}

function sendReminderEmail(){

   
     Logger.log("END sendReminderEmail");
     
}



function checkOverdue(){

      var ss = SpreadsheetApp.getActiveSpreadsheet(); //load the right sheet
      var settingsSheet = ss.getSheetByName("Settings");
      
       var checkboxcontent = settingsSheet.getRange(9,6).getValues();
       
      
      Logger.log("send ovderdue email ? " + checkboxcontent);
      
      if (checkboxcontent == "true"){
      
      
         sendOverDueNotification();
      
      
     } else {
     
       Logger.log("Not sending overdue notices. Check your settings tab");
     }
  

}

function sendOverDueNotification( ){
   
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //load the right sheet
     var itemsSheet = ss.getSheetByName("Items");
     var itemsData = itemsSheet.getDataRange().getValues;
     
      var settingsSheet = ss.getSheetByName("Settings");
      
       var subject = settingsSheet.getRange(13,6).getValues();
       var body = settingsSheet.getRange(13,7).getValues();
     
     var numRows = itemsSheet.getLastRow();
     
     var todayMS = Date.now();

      Logger.log("Today is " + todayMS);
    
      var timeDiff = 0;
    
      
     
       var reminders = [];

      for (var i=1; i<=numRows+1; i++){
      
        //go through all the items
        
          if(itemsSheet.getRange(i,4).getValue() == "Checked Out"){
          
              var dueDate = new Date(itemsSheet.getRange(i,8).getValue());
              var dueDateInMS = dueDate.getTime();
              
              var currentUser = getUserEmailFromID(itemsSheet.getRange(i,5).getValue());
              var currentItem = itemsSheet.getRange(i,2).getValue();
              
              Logger.log("Current User is : " + currentUser);
              
              timeDiff = dueDateInMS - todayMS;
              
              Logger.log("due date in MS " + dueDateInMS);
              Logger.log("Time Diff in MS " + timeDiff);
              
          //if a an item is checked out, add it to the reminders array
          //but only if the time difference is negative
           
              if (timeDiff < 0 ){
              
                //create a new object
                
                var reminder = { 
                  'user': currentUser,
                  'item': currentItem,
                  'dueDate': dueDate
                
                    };
              
                //push it into the array
                
                reminders.push(reminder);
              }
          
          }
      
      }
   Logger.log("Here are the reminders ---------->" );
   
   
   for(var j = 0 ; j<reminders.length ; j++){
   
     Logger.log("reminder number: " + j);
     Logger.log(reminders[j].user);
     Logger.log(reminders[j].item);
     Logger.log(reminders[j].dueDate);
     
     var dateToSend = dueDate.getMonth() +"/"+ dueDate.getDate() +"/"+dueDate.getFullYear();
    
      MailApp.sendEmail(reminders[j].user, 
                   subject,
                   "Hello,  " + "\n" + body + " \n \n ITEM: \n" + reminders[j].item + " \n Which was due on : " + dateToSend );
        
   
   }
   
  
    Logger.log("End sendOverdueNotices ---]");
   
}

function returnItems(returns){
//this function accepts an array of returned materials (from scanned barcode) and 
//updates the location, status, lastUserID, and resets checkout and dueDate


  var ss = SpreadsheetApp.getActiveSpreadsheet(); //load the right sheet
  var itemsSheet = ss.getSheetByName("Items");
  var itemsData = itemsSheet.getDataRange().getValues(); //grab all the data
  
  var settingsSheet = ss.getSheetByName("Settings");
  var libraryLocation = settingsSheet.getRange(4, 5).getValue();
  
 
  
  var statusCell = itemsSheet.getRange(1, 3); 
   
    
  var locationCell = itemsSheet.getRange(1, 5); 
  var lastUserCell = itemsSheet.getRange(1, 4); 
  var checkedOutDateCell = itemsSheet.getRange(1, 6); 
  var dueDateCell = itemsSheet.getRange(1, 7); 
  
   var numRows = itemsSheet.getLastRow(); //need this to index

   //check to see if only one item is here or many, one item is determeind by having an empty array
  if (returns[0][0] == null){
  
   Logger.log("Just one item! : " + returns[0]);
   
   
    for (var j = 0 ; j< numRows ; j++){
          
          
             //go through each row and find a match
             
             if(itemsData[j][0] == returns[0] ){
      
                statusCell = itemsSheet.getRange(j+1, 4); 
                statusCell.setValue("in");
                 //set the new location
                locationCell = itemsSheet.getRange(j+1, 6);
                locationCell.setValue(libraryLocation);
                
                //set the last user
                 
                //update checked out date
                checkedOutDateCell = itemsSheet.getRange(j+1, 7);
                checkedOutDateCell.setValue("");
                
               
                //update due date
                dueDateCell = itemsSheet.getRange(j+1, 8);
                dueDateCell.setValue("");
              
              
              
             
             } else{
              Logger.log("No Match Found");
             
             }
  
    }
  
  } else {
  
  
  Logger.log("More than one item to be returned ");
  
       var returnsArray = returns[0].split(/\r?\n/);
   
      Logger.log("Returning : " + returnsArray);
  
      for( var i = 0 ; i<returnsArray.length ; i++){
      
          Logger.log("Item to be returned : " + returnsArray[i]);
        
          //now access items tab and iterate through rows to match item ID with existing IDs
          
          for (var j = 0 ; j< numRows ; j++){
          
          
             //go through each row and find a match
             
             if(itemsData[j][0] == returnsArray[i] ){
       
              Logger.log("We have a match!! Item# " + returnsArray[i] + " == " + itemsData[j][0]);
              
              
              //update row with new status
                
              
                statusCell = itemsSheet.getRange(j+1, 4); 
                statusCell.setValue("in");
                 //set the new location
                locationCell = itemsSheet.getRange(j+1, 6);
                locationCell.setValue(libraryLocation);
                
                //set the last user
                 
                //update checked out date
                checkedOutDateCell = itemsSheet.getRange(j+1, 7);
                checkedOutDateCell.setValue("");
                
               
                //update due date
                dueDateCell = itemsSheet.getRange(j+1, 8);
                dueDateCell.setValue("");
              
             
             } else{
              Logger.log("No Match Found");
             
             }
          
          
          }
       
        
      }
  
  }

 
Logger.log("End of returnItems");
 

}



function sync() {
  var tasks = FusionTables.Task.list(TABLE_ID);
  // Only run if there are no outstanding deletions or schema changes.
  if (tasks.totalItems == 0) {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); //load the right sheet
  var mapsSheet = ss.getSheetByName("Map");
  //var data = mapsSheet.getDataRange().getValues(); //grab all the data
  
  //var lastCol = mapsSheet.getLastColumn();
  //var lastRow = mapsSheet.getLastRow();
    
    
   var wholeSheet = mapsSheet.getRange('A1:H318');
    var values = wholeSheet.getValues();
    
    if (values.length > 1) {
      var csvBlob = Utilities.newBlob(convertToCsv_(values),
          'application/octet-stream');
      FusionTables.Table.replaceRows(TABLE_ID, csvBlob,
         { isStrict: REQUIRE_SAME_COLUMNS, startLine: FIRST_DATA_ROW - 1 });
      Logger.log('Replaced ' + values.length + ' rows');
    }
  } else {
    Logger.log('Skipping row replacement because of ' + tasks.totalItems +
        ' active background task(s)');
  }
}


/**
 * Converts the spreadsheet values to a CSV string.
 * @param {Array} data The spreadsheet values.
 * @return {string} The CSV string.
 */
function convertToCsv_(data) {
  // See https://developers.google.com/apps-script/articles/docslist_tutorial#section3
  var csv = '';
  for (var row = 0; row < data.length; row++) {
    for (var col = 0; col < data[row].length; col++) {
      var value = data[row][col].toString();
      if (value.indexOf(',') != -1 ||
          value.indexOf('\n') != -1 ||
          value.indexOf('"') != -1) {
        // Double-quote values with commas, double quotes, or newlines
        value = '"' + value.replace(/"/g, '""') + '"';
        data[row][col] = value;
      }
    }
    // Join each row's columns and add a carriage return to end of each row
    // except the last
    if (row < data.length - 1) {
      csv += data[row].join(',') + '\r\n';
    }
    else {
      csv += data[row];
    }
  }
  return csv;
}

function getUserEmailFromID(id){

  

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("Clients");
  var data = userSheet.getDataRange().getValues();
   
  //determine if user already in the system
  //look through the column and find user by email...
  
  var numRows = userSheet.getLastRow(); // returns the index of last row
  
  var userEmail = 'no email registered for this user';
  
    for (var i=1; i<numRows; i++){
      
           if(data[i][USER_ID_COLUMN] == id ){
          
           //if in the system, set userID as currentID
           Logger.log("User id: " + data[i][USER_ID_COLUMN] + " Matches");
           
            userEmail = data[i][USER_EMAIL_COLUMN];
            
            return userEmail;
           
          } 
        
        } //end for
  
    Logger.log("No email found for this user ID ");
    
    return userEmail;

}
