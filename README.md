# OCDSB-Tech-Tokens

Code for sending an email when a user completes a tech token for the OCDSB: https://docs.google.com/spreadsheets/d/1QQ2yQJy4mDPt4H5Wz6VZ1HYRDjo0fP5T2NnL99l57Bw/edit#gid=901404802

Here's how it works:

- Triggers when a teacher submits a response on the attached Google Form. Grabs the form response in an array.
- Checks whether the tab exists for that tool. If it does, it will then search for the teacher and add a checkbox in the appropriate column (for the correct token).
  - At each step of the way, if the tool, pathway, or sheet is missing, it will create it.
- Look in the 'Tokens' tab and select the ID for the appropriate token.
- Send an email to that user with the token as an attachment.

## Manual requirements

- In order to capture the user's name, the AdminSDK needs to be enabled in Google Scripts.
- A manual trigger needs to be added which triggers from the spreadsheet on form submit.
