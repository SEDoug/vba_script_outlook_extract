# vba_script_outlook_extract
This is a VBA script written with Excel | Visual Basic for Applications (VBA).<br>
I created this VB Script on top of a script I found to automate some taks while seeking full stack development knowledge.<br>
Created by @DougMacGregor || http://doug-macgregor.webflow.io/<br>
Seeking a broader field and desire to do work in full stack development.

<hr>

1.	If you don't see the Developer tab, do this first and put a check into the Developer Tab below. Go to File, Options, Customize Ribbon.<br>

![java-code](https://raw.githubusercontent.com/SEDoug/vba_script_outlook_extract/master/img/excel_developer_options.JPG)<br>

Click on Developer, Insert, Active X Controls, command button. After this is setup, save your excel file as an Excel Marco-Enabled Worksheet.
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
3. You need to have Outlook open when you execute the macro from the excel workbook.
<hr>

