VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "btm_extractemails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub btn_extractemails_Click()

Dim OlApp As Object
Dim OlMail As Object
Dim OlItems As Object
Dim OlFolder As Object
Dim j As Integer
Dim strFolder As String

On Error Resume Next
Set OlApp = GetObject(, "Outlook.Application")

If Err.Number = 429 Then
    Set OlApp = CreateObject("Outlook.Application")
End If

strFolder = ThisWorkbook.Path & "\Reports_Emailed" 'Change Reports_Emailed to the folder name where you would like to have your attachements saved

Set OlFolder = OlApp.getnamespace("MAPI").Folders("NBS IT Reports").Folders("Inbox").Folders("User Access Reviews").Folders("201904")
'Change NBS IT Reports to your Main mailbox
'Change Inbox to your sub folder of the Main mail box, etc...

Set OlItems = OlFolder.Items

For Each OlMail In OlItems
    If OlMail.Attachments.Count > 0 Then
        For j = 1 To OlMail.Attachments.Count
            OlMail.Attachments.Item(j).SaveAsFile strFolder & "\" & OlMail.Attachments.Item(j).Filename
        Next j
    End If
Next

Set OlFolder = Nothing
Set OlItems = Nothing
Set OlMail = Nothing
Set OlApp = Nothing

MsgBox "Done", vbInformation

End Sub

