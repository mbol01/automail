# Automail
Automatically send an email to form respondants depedning if they got a spot to an activity or not
## Background
This is a google appscript designed to automaticaly send an email to Google Forms respondants depending on if the arrangement they applied for had spots left. If no spots are left they will automatically recieve an email saying they are on the reserve list. The script was originally designed to be used during the reception of new students to Automation and mechatronics at Chalmers in 2023 and some variables and all emails are therefor written in Swedish but with basic HTML-programming skills this should easily be adopted to your needs.

## Setting up
This guide will help ypu setup and start using the script. Make sure you have a google account to start with.
### Preparing the sheet
Download template.xlsx to your computer and upload it to your Google account. Open it and press File -> Save as Google Sheets. You can now remove the .xlsx from your drive. Rename the file from template to your desired name.

### Setting up the code
Go to Extensions -> App Script and name your project. Next, copy the contents of code.gs and all the .html files in this repo to corresponding files in the Files section. Unfortunately Google App Scripts seems to have no way to upload files. Replace all stubs ([...]) with your own variables, the Sheet-ID can be found in the link of your sheet as below:
```
https://docs.google.com/spreadsheets/d/this_is_your_sheet_id/edit 
```
Also make sure to edit the .html files so that they match the information you want included. Text surrounded by `?=  ?>` are variables that will be changed from form to form using the config sheet.

### Setting up the trigger
To have the code triggered by submitted form you need to setup a trigger. Go to the "Trigger" tab in App Script and press "Add Trigger" in the bottom right corner.
Choose "onFormSubmit" as the function, "Head" as the deployment, "From spreadsheet" as the event source and "On form submit" as event type. In the pop-up-window, log in, press see advanced and proceed. Allow the script the necessary actions.

### Create form template
Unfortunately, Google does not provide a way to share a form outside their platform. Therefore you have to manually create a template with the questions the form uses. You can of course change these questions and their order but note that these changes needs to be made in code.gs and the spreadsheet as well! The form should be setup to collect emails.
These are the questions in order:
1. Name
2. Confirm email (some people are in a hurry and having the respondants entering it twice increases the chances of one being correct)
3. Food preference (multiple choice is easier to read later but not required for the script to work)
4. Other
5. GDPR

More questions can be added between Q3 and Q4.

## Create a form and use the script
Now we come to the fun part! 
### Clone your template form
Take your template form, clone it and change the name and description to something relevant. Also add any extra questions you need.
### Linking form to your spreadsheet
Press the "Responses"-tab in your cloned form and press Link to Sheets -> Select existing spreadsheet and choose your spreadsheet from earlier. This will now open that spreadsheet. Rename the sheet created by the form (which will have a purple logo indicated it's linked to a form) to something unique, usually the name of the activity. Add this name (case-sensitive) to the name column of the "Config"-sheet and enter the cost, the discounted cost, the cost of prepurchased alcohol (leave blank if not applicable), number of spots and information for ticket sales in their respective column. If the cost cell is left blank the activity will be considered free and that email template will be used instead. It's very important that the sheet-name for the form-sheet and the name-cell in "Config" precisely match each other. 

### Done
The form respondants should now get emails from the script!
A log can be found in the Executions tab of App Script.

## Prepare for ticketsale function
The script will add a button called "Infochef" in the Sheet menu. When pressed, it gives the option to prepare for sale and the option to reset from it. The prepare for sale checks all emails and if they are present in column A in "Mottagningsavgift" sheet their rows will be marked orange. This is to make it easier to handle discounts. It also hides irrelevant rows and sets a column heading to "Paid" and adds counters to the last cell in that row to make it easier to keep track on how many tickets are left. The reset button only unhides the hidden columns.

## Known bugs and limitations
Due to the code being sent by a script it will get caught in some spam-filters.