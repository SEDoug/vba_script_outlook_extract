# vba_script_outlook_extract
This is a VBA script written using with Excel | Visual Basic for Applications (VBA).<br>
I created this VB Script on top of a script I found to showcase skills while seeking full stack development knowledge.<br>
Created by @DougMacGregor || http://doug-macgregor.webflow.io/<br>
Seeking a broader field and desire to do work in full stack development.

<hr>

1.	If you don't see the Developer tab, do this first and put a check into the Developer Tab below. Go to File, Options, Customize Ribbon.<br>

![java-code](https://raw.githubusercontent.com/SEDoug/vba_script_outlook_extract/master/img/excel_developer_options.JPG)<br>

Click on Developer, Insert, Active X Controls, command button. After this is setup, save your excel file as an Excel Marco-Enabled Worksheet.
![java-code](https://raw.githubusercontent.com/SEDoug/vba_script_outlook_extract/master/img/excel_developer_activeX.PNG)<hr>

2.	VBA Script
a.	We write the explicit option to keep all variables isolated.
b.	The next few lines we are creating our Outlook objects, integers, strings.
c.	Next we need to check if our Outlook application is open. If not open then the script will open a new Outlook application window.
d.	We define where we want to extract the attachments and create a Folder to save the attachments.
e.	Now we set which Outlook Main mailbox and sub mails box we are working with.
f.	We now set the email to a specific object.  
g.	Now we loop through all emails for attachments in the specified mailbox.
h.	As the script finds an attachment, it saves the attachment into the specified folder. This keeps looping until no other emails are found.
i.	Finish with setting the objects to nothing after the loop.
j.	Now finish with a “Done” message after the script is finished looping.

![java-code](https://raw.githubusercontent.com/SEDoug/vba_script_outlook_extract/master/img/vba_script_02.PNG)<hr>
3. You need to have Outlook open when you execute the macro from the excel workbook.
<hr>

