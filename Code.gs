function onSubmit(e) {
  const formResponses = e.values;
  const row = e.range.rowStart;
  const newArr = formResponses.filter((el) => el);
  const ss = SpreadsheetApp.getActive();
  
  
  const email = newArr[1].trim();
  const tool = newArr[2].trim();
  const path = movePathway(row, formResponses);
  // Find the sheet for the tool. If tool not found, add a new tab and style it with some defaults.
  let toolSheet = ss.getSheetByName(tool);
  if (!toolSheet) {
    toolSheet = ss.insertSheet().setName(tool);
    toolSheet.getRange("A1").setValue("Teacher");
    toolSheet.setFrozenColumns(1);
    toolSheet.setFrozenRows(1);
    };

  // Find the person's email in the Tool Tab. If not found, add them to the end.
  let userRow = toolSheet.getRange(2,1,toolSheet.getLastRow()).getValues().findIndex((user) => user[0].trim() === email) + 2;
  if (userRow < 2) {
    userRow = toolSheet.getLastRow()+1;
    toolSheet.getRange(userRow,1).setValue(email)
    };

  // Find the pathway beside the person. If not found, add the pathway.
  let pathCol = toolSheet.getRange(1,2,1,toolSheet.getLastColumn()).getValues()[0].findIndex((item) => item.trim() === path) + 2;
  if (pathCol < 2) {
    pathCol = toolSheet.getLastColumn()+1;
    toolSheet.getRange(1,pathCol).setValue(path)
    };

  // Add a checkbox beside the person's email in the appropriate column.
  toolSheet.getRange(userRow,pathCol).insertCheckboxes().check();

  // Get the ID to attach to the email.
  let tokenID, pathway;
  try {
    tokenID = ss.getSheetByName('Tokens').getDataRange().getValues().find((token) => token[0].toString().includes(`${tool} - ${path}`))[1];
    pathway = `${tool} - ${path}`;
  } catch (e) {
    tokenID = ss.getSheetByName('Tokens').getDataRange().getValues().find((token) => token[0].toString().includes(tool))[1];
    pathway = tool;
  }

  // Find the person's first name.
  let name;
  try {
    name = AdminDirectory.Users.get(email, {viewType:'domain_public', fields:'name'}).name.givenName;
  } catch(e) {
    name = capitalizeString(email.split('.')[0]);
  }

  // Send email with the token as an attachment.
  MailApp.sendEmail({
    to: email,
    subject: `Congratulations on completing the ${pathway} Tech Token Learning Path!`,
    htmlBody:`<p>Hi ${name}!</p>
    <p>Way to go! You've conquered the ${pathway} Pathway, and we're thrilled to reward you with the attached token. Want to show off your accomplishment? Check out <a href="https://drive.google.com/file/d/1CawdWcyvUzLDIR5Jx2eG0KMDX_SQC4ga/view?usp=sharing" target="_blank">these instructions</a> to add it to your signature. 
Pathways will continue to be added, so check back for new Tech Tokens to add to your collection. Keep shining, tech wizard!</p>

<p>The B&LT Academic Team</p>
    `,
    attachments: DriveApp.getFileById(tokenID),
    name: "B&LT Academic Team",
  });
}

function capitalizeString(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}

function movePathway(row, formResponse) {
  try {
  const ss = SpreadsheetApp.getActive().getActiveSheet();
  const colHeaders = ss.getDataRange().getValues()[0];
  const firstPathCol = colHeaders.findIndex((el) => el.toString().toLowerCase().includes("which pathway")) + 1;

  const {index, path} = formResponse.map((el, i) => {
    if (el && colHeaders[i].toString().toLowerCase().includes("which pathway")) {
      return {index: i+1, path: el};
    }
  }).find((el) => el);

  ss.hideColumns(index);
  ss.getRange(row, index).clear();
  ss.getRange(row, firstPathCol).setValue(path);

  return path;
  } catch (e) {
    return ""
  }
}
