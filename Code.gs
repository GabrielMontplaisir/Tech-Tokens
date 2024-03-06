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
  
  const email = newArr[1].trim();                     // Save the user's email separately. The .trim() will eliminate all empty spaces before and after the email address.
  const tool = newArr[2].trim();                      // Save the tool separately. Eg. Canva, Minecraft, WeVideo.
  const path = movePathway(row, formResponses);       // Calls movePathway(). Takes in the row and form responses. You can find the movePathway() function towards the bottom of this page.
  const pathway = (!path) ? tool : `${tool} - ${path}`; // Set some formatting for the email.

  // Find the person's first name. Adds a little "personal" touch to the email.
  let name;
  try {
    // Here, we're accessing public Google information related to the person's profile. Set by the OCDSB domain.
    name = AdminDirectory.Users.get(email, {viewType:'domain_public', fields:'name'}).name.givenName;
  } catch(e) {
    // If the above failed, use the prefix in the person's email. Retrieve the part before the period, and capitalize the first letter using capitalizeString() function. See below for that code.
    name = capitalizeString(email.split('.')[0]);
  }

  /* Call on updateToolSheet() to update the tool's sheet tab with the teacher email, the pathway and a checkbox beside the teacher's name.
  *  Returns the number of tokens the teacher has for that tool.
  */
  const teacher = updateToolSheet(email, tool, path);
  const teacherNumTokens = teacher.numTokens;
  const teacherExpert = teacher.expert;
  const completed = teacher.completed;

  /* This code segment calls on findToken() to find the Tech token to send to the user. It also returns how many tech tokens exist for the tool, and some formatting for the pathway.
  *  If it fails, it will send an error email.
  *  See findToken() for more details.
  */
  let tokenFile, toolNumTokens, expertTokenFile;
  try {
    const token = findToken("Tokens", tool, path, pathway);
    tokenFile = token.tokenFile;
    toolNumTokens = token.numTokens;
  } catch(err) {
    emailError("regular", email, tool, name);
    return;
  }

  /* Send email with the token as an attachment.
  *  Anything contained within ${} is a dynamic value based on the above.
  *  The email body is coded in HTML.
  */
  if (!completed) {
    MailApp.sendEmail({
      to: email,
      subject: `Congratulations on completing the ${pathway} Tech Token Learning Path!`,
      htmlBody:`<p>Hi ${name},</p>
      <p>Way to go! You've conquered the ${pathway} Pathway, and we're thrilled to reward you with the attached token. Want to show off your accomplishment? Check out <a href="https://drive.google.com/file/d/1CawdWcyvUzLDIR5Jx2eG0KMDX_SQC4ga/view?usp=sharing" target="_blank">these instructions</a> to add it to your signature. Pathways will continue to be added, so check back for new Tech Tokens to add to your collection. Keep shining, tech wizard!</p>
      
      <p>The B&LT Academic Team</p>`,
      attachments: tokenFile,
      name: "B&LT Academic Team",
    });
  }


  /*  Send a second email with an Expert Token if the teacher has not received one yet, and if they meet or exceed the number of tech tokens found in the "Tokens" tab.
  *   You can see this number requirement in the "Expert Tokens" tab.
  */
  if (teacherNumTokens >= toolNumTokens && !teacherExpert) {
    try {
      const expertToken = findToken("Expert Tokens", tool);

      if (!expertToken) return;
      expertTokenFile = expertToken.tokenFile;

      /* Send email with the expert token as an attachment.
      *  Anything contained within ${} is a dynamic value based on the above.
      *  The email body is coded in HTML.
      */
      MailApp.sendEmail({
        to: email,
        subject: `You've earned an Expert Token for ${tool}!`,
        htmlBody:`<p>Hi ${name},</p>
        <p>We noticed that you've completed ALL Tech Token Learning Pathways for ${tool}. What an achievement, and a whole lot of learning! Feel free to share everything you've learned with other educators at your site. Feel free to replace any current tokens for ${tool} you have in your signature with this sparkly new one. Additional pathways will continue to be added in the future, so check back for new Tech Tokens to add to your collection. Keep shining, tech wizard!</p>
        
        <p>The B&LT Academic Team</p>`,
        attachments: expertTokenFile,
        name: "B&LT Academic Team",
      });

    } catch(err) {
      emailError("expert", email, tool, name);
    }

    // Update Expert column with a checkbox.
    updateToolSheet(email, tool, "Expert");
  }
}


// A simple function to capitalize a word. JavaScript does not do this by default.
function capitalizeString(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}

// Function to send error email to user if there is an error.
function emailError(tokenType, email, tool, name) {
  let subjectLine, emailBody;
  if (tokenType === "regular") {
    subjectLine = `Congratulations on completing the ${tool} Tech Token Learning Path!`; 
    emailBody = `
      <p>Hi ${name}!</p>
      <p>Way to go! You've conquered the ${tool} Pathway! Pathways will continue to be added, so check back for new Tech Tokens to add to your collection. Unfortunately, there was a slight problem on our end and we could not automatically retrieve the tech token for you. If you would like to receive the tech token for your completion, please email <a href="mailto:blt.academic@ocdsb.ca">blt.academic@ocdsb.ca</a>, and we will send it to you as soon as possible!</p>
      
      <p>Keep shining, tech wizard!</p>
      
      <p>The B&LT Academic Team</p>
    `;
  } else {
    subjectLine = `You've earned an Expert Token for ${tool}!`; 
    emailBody = `
      <p>Hi ${name},</p>
        <p>We noticed that you've completed ALL Tech Token Learning Pathways for ${tool}. What an achievement, and a whole lot of learning! Feel free to share everything you've learned with other educators at your site. Pathways will continue to be added, so check back for new Tech Tokens to add to your collection. Unfortunately, there was a slight problem on our end and we could not automatically retrieve the expert token for you. If you would like to receive it, please email <a href="mailto:blt.academic@ocdsb.ca">blt.academic@ocdsb.ca</a>, and we will send it to you as soon as possible! Once you do, feel free to replace any current tokens for ${tool} you have in your signature with this sparkly new one.</p>
      
      <p>Keep shining, tech wizard!</p>
      
      <p>The B&LT Academic Team</p>
    `;
  }

  MailApp.sendEmail({
    to: email,
    bcc: "blt.academic@ocdsb.ca",
    subject: subjectLine,
    htmlBody: emailBody,
    name: "B&LT Academic Team",
  });
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

    if (index !== firstPathCol) ss.hideColumns(index);  // For cleanliness, we hide the original question because we're moving the response to the firstPathCol, unless it's the firstPathCol.
    ss.getRange(row, index).clear();                  // Delete the original response.
    ss.getRange(row, firstPathCol).setValue(path);    // Place the original response to the firstPathCol in the correct row.

    // Return the name of the pathway (string)
    return path;
  } catch (e) {
    // If there's an error in the code above, then return an empty string.
    return ""
  }
}


function findToken(tab, tool, path, pathway) {
  const ss = SpreadsheetApp.getActive();
  const tokenList = ss.getSheetByName(tab).getDataRange().getValues();
  const listTokensForTool = tokenList.filter((token) => token[0].trim().toLowerCase() === tool.toLowerCase());

  if (listTokensForTool.length < 1) {
    if (tab === "Expert Tokens") return false;
    throw(`The tool (${tool}) could not be found in the Tokens tab. Please ensure that it exists and its name matches the one found in the Google Form, along with its pathway and TokenID.`);
  }

  let tokenID;
  try {
    if (tab === "Tokens"){
      tokenID = listTokensForTool.find((token) => token[1].trim().toLowerCase() === path.toLowerCase())[2].toString().trim();
    } else {
      tokenID = listTokensForTool[0][2];
    };
  } catch (err) {
    throw(`The pathway (${path}) for ${tool} does not exist in the Tokens tab. Please ensure it exists and its name matches the one found in the Google Form, along with its TokenID.`);
  }

  console.log(tokenID);

  let tokenFile;
  try {
    if (!tokenID) {
      throw err;
    };
    tokenFile = DriveApp.getFileById(tokenID); 
  } catch (err) {
    throw(`Could not find ID for ${pathway} in Tokens tab. Please ensure the ID is present in column C, spelled properly, and BLT Academic has share permissions for it.`);
  }

  return {tokenFile, numTokens: listTokensForTool.length};                                                                                      
}

function updateToolSheet(email, tool, path) {
  const ss = SpreadsheetApp.getActive();              // The Form Responses tab in the spreadsheet.

  // Find the sheet for the tool. If tool not found, add a new tab and style it with some defaults.
  let toolSheet = ss.getSheetByName(tool);            // Retrieve the tab called <tool>. Eg. Canva, ThingLink. Save it to its own variable.
  if (!toolSheet) {                                   // The exclamation point is another way to say NOT. In this case: IF (NO TOOLSHEET) THEN ...
    toolSheet = ss.insertSheet().setName(tool);       // Insert a sheet and set its name to <tool>.
    toolSheet.getRange("A1").setValue("Teacher");     // Set some preset default formatting to the sheet. Set cell A1 to "Teacher".
    toolSheet.getRange("B1").setValue("Expert");     // Set some preset default formatting to the sheet. Set cell A1 to "Teacher".
    toolSheet.hideColumns(2);
    toolSheet.setFrozenColumns(1);                    // Freeze the first column.
    toolSheet.setFrozenRows(1);                       // Freeze the first row.
  };


  /* Find the person's email in the tool's tab. If not found, add them to the end.
  *  In the tool sheet, grab all values in the Teacher column and try to find the respondent's email.
  *  We add + 1 because arrays start at 0, but rows in Google Sheets start at 1.
  *  In the event a user is NOT found, a findIndex() function returns -1. Because we add +1 though, our "error" returns as 0; hence why we check if userRow < 1.
  */
  let user = toolSheet.getDataRange().getValues();
  let userRow = user.findIndex((user) => user[0].trim() === email) + 1;
  if (userRow < 1) { 
    userRow = toolSheet.getLastRow()+1;                // Set the userRow to the last row + 1 because we'll be adding a new row.
    toolSheet.getRange(userRow,1).setValue(email);     // Add the teacher to this spot.
    toolSheet.getRange(userRow, 2).insertCheckboxes(); // Insert a checkbox in the Expert column.
  };

  /* Find the pathway beside the person in the tool's tab. If not found, add the pathway at the end.
  *  This segment uses the same logic as the userRow, but for the pathway instead.
  */
  let pathCol = user[0].findIndex((item) => item.trim() === path) + 1;
  if (pathCol < 1) {
    pathCol = toolSheet.getLastColumn()+1;
    toolSheet.getRange(1,pathCol).setValue(path);
  };

  const rowCol = toolSheet.getRange(userRow,pathCol);

  const completed = rowCol.getValue();
  // Add a checkbox in the appropriate userRow and pathCol. Check the box.
  rowCol.insertCheckboxes().check();
  
  // The first index is their email. Therefore, the number of tokens the user has is the array length - 1.
  user = toolSheet.getDataRange().getValues(); // Reinitiate user.
  const numTokens = user.find((value) => value[0].trim() === email).filter((el) => el === true).length;

  return {numTokens, completed, expert: toolSheet.getRange(userRow, 2).getValue()}
}

