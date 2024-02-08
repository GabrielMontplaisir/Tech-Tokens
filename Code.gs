/* This script automatically retrieves form responses on submission from https://docs.google.com/forms/d/1AUYyGcqMSBNeHG17gOVa4NcP-Cd5kYH33dU_q8MVKUc/edit
*  and sends an email to the responder with a pre-created message along with an image of the tech token.
*  I will add comments throughout to explain what everything does to the best of my abilities.
*/

// For Form Submit Events to work, it must be triggered via an Event Trigger. On the left hand side, look at the "Triggers" tab, created by the B&LT Academic generic account. It checks for Form submissions before triggering this script.

/* @param {Event} e The Form Submit event. More details here (Under Google Sheets Events > Form Submit): https://developers.google.com/apps-script/guides/triggers/events
*  The "event" (e) will output the form values, kind of how it's spit out in the Google Sheet in a typical "Form Responses" tab. We manipulate this data as we need.
*/

function onSubmit(e) {
  const formResponses = e.values;                     // The "raw" form response ansers from the user. Includes the user email.
  const row = e.range.rowStart;                       // The row where the responses are placed in the spreadsheet.
  const newArr = formResponses.filter((el) => el);    // By default, not all columns in the "Form Responses" tab are filled, we therefore need to eliminate all empty values from the formResponses.
  const ss = SpreadsheetApp.getActive();              // The Form Responses tab in the spreadsheet.
  
  
  const email = newArr[1].trim();                     // Save the user's email separately. This helps with the readability. The .trim() will eliminate all empty spaces before and after the email address.
  const tool = newArr[2].trim();                      // Save the tool separately. Eg. Canva, Minecraft, WeVideo.
  const path = movePathway(row, formResponses);       // This calls the function movePathway(). Takes the row and form responses as parameters. You can find the movePathway() function towards the bottom of this page.

  // Find the sheet for the tool. If tool not found, add a new tab and style it with some defaults.
  let toolSheet = ss.getSheetByName(tool);            // Retrieve the tab called <tool>. Eg. Canva, ThingLink. Save it to its own variable.
  if (!toolSheet) {                                   // The exclamation point is another way to say NOT. In this case: IF (NO TOOLSHEET) THEN ...
    toolSheet = ss.insertSheet().setName(tool);       // Insert a sheet and set its name to <tool>.
    toolSheet.getRange("A1").setValue("Teacher");     // Set some preset default formatting to the sheet. Set cell A1 to "Teacher".
    toolSheet.setFrozenColumns(1);                    // Freeze the first column.
    toolSheet.setFrozenRows(1);                       // Freeze the first row.
  };


  // Find the person's email in the Tool Tab. If not found, add them to the end.
  let userRow = toolSheet.getRange(2,1,toolSheet.getLastRow()).getValues().findIndex((user) => user[0].trim() === email) + 2;   // In the tool sheet, we're collecting all values in the Teacher column and trying to find the responder.
  if (userRow < 2) {                                  // If the above doesn't find the person's email, it will return -1. However, we added +2, meaning that our "error" is actually = 1. We have to check if the userRow is < 2.
    userRow = toolSheet.getLastRow()+1;               // Set the userRow to the last row + 1 because we'll be adding a new row.
    toolSheet.getRange(userRow,1).setValue(email)     // Add the teacher to this spot.
  };

  // Find the pathway beside the person. If not found, add the pathway.
  // This little segment uses the same logic as the userRow, but for the pathway instead.
  // This will add the pathway if not found to the last column in the toolSheet.
  let pathCol = toolSheet.getRange(1,2,1,toolSheet.getLastColumn()).getValues()[0].findIndex((item) => item.trim() === path) + 2;
  if (pathCol < 2) {
    pathCol = toolSheet.getLastColumn()+1;
    toolSheet.getRange(1,pathCol).setValue(path)
  };

  // Add a checkbox in the appropriate userRow and pathCol. Check the box.
  toolSheet.getRange(userRow,pathCol).insertCheckboxes().check();



  // Find the Token for the right tool and pathway.
  let tokenID, pathway;

  // If the "path" from the movePathway() returned an empty string, then we know the following will "fail".
  try {
    tokenID = ss.getSheetByName('Tokens').getDataRange().getValues().find((token) => token[0].toString().includes(`${tool} - ${path}`))[1];     // Get the ID from the "Tokens" tab if it matches to "<tool> - <path>".
    pathway = `${tool} - ${path}`;                                                                                                              // Save the format for simplicity in our email.
  } catch (e) {
    tokenID = ss.getSheetByName('Tokens').getDataRange().getValues().find((token) => token[0].toString().includes(tool))[1];                    // Since it failed, all we're looking for is the tool instead.
    pathway = tool;                                                                                                                             // Save the format for simplicity in our email.
  }

  // Find the person's first name. Adds a little "personal" touch to the email.
  let name;
  try {
    name = AdminDirectory.Users.get(email, {viewType:'domain_public', fields:'name'}).name.givenName;       // Here, we're accessing public Google information related to the person's profile. Set by the OCDSB domain.
  } catch(e) {
    name = capitalizeString(email.split('.')[0]);     // If the above failed, then we're going to use the prefix in the person's email. Retrieve the part before the period, and capitalize the first letter using capitalizeString() function. See below for that code.
  }

  // Send email with the token as an attachment.
  // Anything contained within ${} is a dynamic value based on the above.
  // The email body is coded in HTML.

  MailApp.sendEmail({
    to: email,
    subject: `Congratulations on completing the ${pathway} Tech Token Learning Path!`,
    htmlBody:`<p>Hi ${name}!</p>
    <p>Way to go! You've conquered the ${pathway} Pathway, and we're thrilled to reward you with the attached token. Want to show off your accomplishment? Check out <a href="https://drive.google.com/file/d/1CawdWcyvUzLDIR5Jx2eG0KMDX_SQC4ga/view?usp=sharing" target="_blank">these instructions</a> to add it to your signature. Pathways will continue to be added, so check back for new Tech Tokens to add to your collection. Keep shining, tech wizard!</p>
    
    <p>The B&LT Academic Team</p>`,
    attachments: DriveApp.getFileById(tokenID),
    name: "B&LT Academic Team",
  });
}


// A simple function to capitalize a word. JavaScript does not do this by default.
function capitalizeString(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}


/* A function which takes the form responses, finds the response associated to the pathway and moves it to the first column for pathway.
*  This function is needed because Google Forms creates a new column EVERY time a new question is created for a pathway. Eg. the "pathway" column for Canva is not the same "pathway" column for Micro:bits.
*  This helps with readability, but also helps the script stay "consistent" in finding the pathway.
*
*  If the function "fails" because it can't find the pathway (ThingLink does NOT have a pathway for example), it will return an empty string. We'll use this empty string later.
*  Try..catch blocks are a way to test code for errors. This is how we'll return an empty string if it doesn't work.
*/

function movePathway(row, formResponse) {
  try {
  const ss = SpreadsheetApp.getActive().getActiveSheet();                                                             // Retrieve the current sheet (Form Responses)
  const colHeaders = ss.getDataRange().getValues()[0];                                                                // Retrieve the column headers for the first row in the sheet.
  const firstPathCol = colHeaders.findIndex((el) => el.toString().toLowerCase().includes("which pathway")) + 1;       // Find the first question which begins with "which pathway". We'll place the pathway from the form response to this column later.


  // Iterate through the raw form responses, and create a new array (a list of items) where the item in the form response corresponds to the "which pathway" question.
  // This should return a singular item.
  // We then find it and assign the index (column) and the pathway (path) to variables. 

  const {index, path} = formResponse.map((el, i) => {
    if (el && colHeaders[i].toString().toLowerCase().includes("which pathway")) {
      return {index: i+1, path: el};
    }
  }).find((el) => el);


  ss.hideColumns(index);                            // For cleanliness, we hide the original question because we're moving the response to the firstPathCol.
  ss.getRange(row, index).clear();                  // Delete the original response.
  ss.getRange(row, firstPathCol).setValue(path);    // Place the original response to the firstPathCol in the correct row.

  return path;                                      // We return the pathway (string) to use later to find the token and use in the email.
  } catch (e) {
    return ""                                       // If anything breaks, then return an empty string.
  }
}
