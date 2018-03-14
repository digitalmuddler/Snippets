Private Sub btnCreateEmail_Click()
   
   Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Have you filled out data on email form? "
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "Email Confirmation"
    Ctxt = 1000
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    If Response = vbNo Then    ' User chose Yes.
        Exit Sub
    End If
   
   ActiveSheet.Range("BA1:BH22").Select 'data to send
      
   ActiveWorkbook.EnvelopeVisible = True   ' Show the envelope on the ActiveWorkbook.
   
   ' Set the optional introduction field thats adds some header text to the email body. It also sets the To, CC, and Subject lines.
   With ActiveSheet.MailEnvelope
      .Introduction = "Insert additional information and instruction here."
      .Item.To = ActiveSheet.Range("C5") 'automatically adds to TO field
      .Item.cc = ActiveSheet.Range("C4") 'automatically adds name to the CC field
      .Item.Subject = "SUBJECT INFO HERE " & ActiveSheet.Range("C2") & " - " & ActiveSheet.Range("C3") & " " & ActiveSheet.Range("D3") 'pulls info from select cells
   End With
End Sub
