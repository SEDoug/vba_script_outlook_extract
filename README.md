# vba_script_outlook_extract
I rolled up this VBA script to automate a monthly task. With the click of a button all Outlook attachments can be saved to a folder.
Created by @DougMacGregor || http://doug-macgregor.webflow.io/ | Seeking a broader field and desire to do work in full stack development.<br>

A VBA script can be used as a great tool to help improve a monthly task of saving multiple email attachments from any Outlook Inbox.  When I encounter a work task that requires some manual work I look for ways on “how to automate” this.  Clicking multiple emails and saving multiple attachments felt like a mindless task. Luckily for some, we have a thirst for knowledge and I wanted to share what I recently rolled up.<br>

Here’s How It Works | written using with Excel | Visual Basic for Applications (VBA)

# Step-by-step guide
<hr>

If you don't see the Developer tab in Excel, do this first and put a check into the Developer Tab below. Go to File, Options, Customize Ribbon.<br>

![java-code](https://raw.githubusercontent.com/SEDoug/vba_script_outlook_extract/master/img/excel_developer_options.JPG)<br>

After this is setup, save your excel file as an Excel Marco-Enabled Worksheet. Click on Developer, Insert, Active X Controls, command button. Name the command button Extract Emails or what you see fit.<br>

![java-code](https://raw.githubusercontent.com/SEDoug/vba_script_outlook_extract/master/img/excel_developer_activeX.PNG)<hr>

# VBA Script
a.	First, write the explicit option to keep all variables isolated.<br>
b.	Next create all the Outlook objects, integers, strings.<br>
c.	Second, is to check if Outlook is open. If not, then the script will open a new Outlook application window.<br>
d.	Next, define where to extract the attachments, create the folder, and save the attachments.<br>
e.	Now set which Outlook Main mailbox and sub-mail box to work with.<br>
f.	Next set the items object to a specific folder object.<br>
g.	Now loop through all emails for attachments in the specified mailbox.<br>
h.	As the script finds an attachment, it saves the attachment into the specified folder (d), the loop will run through all emails.<br>
i.	Finish with setting the objects to nothing after the loop.<br>
j.	Script completes with a “Done” message after the loop is finished.<br>

![java-code](https://raw.githubusercontent.com/SEDoug/vba_script_outlook_extract/master/img/vba_script_02.PNG)<hr>
3. You need to have Outlook open when you execute the macro from the excel worksheet.
<hr>

