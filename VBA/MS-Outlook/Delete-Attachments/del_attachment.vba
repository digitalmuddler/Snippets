' found this on some random website
' deletes attachments and adds a list of the filenames deleted to the bottom of the message

Sub Delete_attachments_nosave()
 
 Dim Response As VbMsgBoxResult
 Response = MsgBox("Do you REALLY want to PERMANENTLY delete all attachments in all SELECTED mails?" _
 , vbExclamation + vbDefaultButton2 + vbYesNo)
 
 If Response = vbNo Then Exit Sub
 
 Dim myAttachment As Attachment
 Dim myAttachments As Attachments
 Dim selItems As Selection
 Dim myItem As Object
 Dim lngAttachmentCount As Long
 
 ' Set reference to the Selection.
 Set selItems = ActiveExplorer.Selection
 
 ' Loop though each item in the selection.
 For Each myItem In selItems
   Set myAttachments = myItem.Attachments
 
   lngAttachmentCount = myAttachments.Count
 
 ' Loop through attachments until attachment count = 0.
 While lngAttachmentCount > 0
   strFile = myAttachments.Item(1).FileName & "; " & strFile
   myAttachments(1).Delete
   lngAttachmentCount = myAttachments.Count
 Wend
  
  
 If myItem.BodyFormat <> olFormatHTML Then
   myItem.Body = myItem.Body & vbCrLf & _
   "The file(s) removed were: " & strFile
 Else
   myItem.HTMLBody = myItem.HTMLBody & "<p>" & _
   "The file(s) removed were: " & strFile & "</p>"
 End If
 
 myItem.Save
  strFile = ""
 Next
 
 MsgBox "Done. All attachments were deleted.", vbOKOnly, "Message"
 
 Set myAttachment = Nothing
 Set myAttachments = Nothing
 Set selItems = Nothing
 Set myItem = Nothing
End Sub
