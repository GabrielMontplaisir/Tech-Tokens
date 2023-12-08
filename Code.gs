function onSubmit(e) {
  const formResponses = e.values;
  const newArr = formResponses.filter((el) => el);
  const ss = SpreadsheetApp.getActive();
  
  // Find the sheet for the tool. If tool not found, add a new tab and style it with some defaults.
  let toolSheet = ss.getSheetByName(newArr[2]);
  if (!toolSheet) {
    toolSheet = ss.insertSheet().setName(newArr[2]);
    toolSheet.getRange("A1").setValue("Teacher");
    toolSheet.setFrozenColumns(1);
    toolSheet.setFrozenRows(1);
    };

  // Find the person's email in the Tool Tab. If not found, add them to the end.
  let userRow = toolSheet.getRange(2,1,toolSheet.getLastRow()).getValues().findIndex((user) => user[0] === newArr[1]) + 2;
  if (userRow === 1) {
    toolSheet.getRange(toolSheet.getLastRow()+1,1).setValue(newArr[1])
    userRow = toolSheet.getLastRow();
    };

  // Find the pathway beside the person. If not found, add the pathway.
  let pathCol = toolSheet.getRange(1,2,1,toolSheet.getLastColumn()).getValues().findIndex((path) => path[0] === newArr[3]) + 2;
  if (pathCol === 1) {
    toolSheet.getRange(1,toolSheet.getLastColumn()+1).setValue(newArr[3])
    pathCol = toolSheet.getLastColumn();
    };

  // Add a checkbox beside the person's email in the appropriate column.
  toolSheet.getRange(userRow,pathCol).insertCheckboxes().check();

  // Get the ID to attach to the email.
  const tokenID = ss.getSheetByName('Tokens').getDataRange().getValues().find((token) => token[0].toString().includes(newArr[2]))[1];

  // Find the person's first name.
  let name;
  try {
    name = AdminDirectory.Users.get(newArr[1], {viewType:'domain_public', fields:'name'}).name.givenName;
  } catch(e) {
    name = capitalizeString(newArr[1].split('.')[0]);
  }

  // Send email with the token as an attachment.
  MailApp.sendEmail({
    to: newArr[1],
    subject: `Congratulations on completing the ${newArr[2]} - ${newArr[3]} Tech Token Learning Path!`,
    htmlBody:`<p>Hi ${name}!</p>
    <p>Way to go! You've conquered the ${newArr[2]} - ${newArr[3]} Pathway, and we're thrilled to reward you with the attached token. Want to show off your accomplishment? Check out <a href="https://drive.google.com/file/d/1CawdWcyvUzLDIR5Jx2eG0KMDX_SQC4ga/view?usp=sharing" target="_blank">these instructions</a> to add it to your signature. 
Pathways will continue to be added, so check back for new Tech Tokens to add to your collection. Keep shining, tech wizard!</p>

<p>The B&LT Academic Team</p>
    `,
    attachments: DriveApp.getFileById(tokenID)
  });
}

function capitalizeString(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}
