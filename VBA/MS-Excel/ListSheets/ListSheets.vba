'create as Module

Sub ListSheets()
 ' creates list of tabs w/ hyperlinks to each tab 
 
Dim ws As Worksheet
Dim x As Integer
 
 x = 7 'starting row
 
 Sheets("SUMMARY").Range("B7:B500").Clear  'clears location of tab list
 
For Each ws In Worksheets
 'adds hyperlink
 Sheets("SUMMARY").Cells(x, 2).Select 
 ActiveSheet.Hyperlinks.Add _
 Anchor:=Selection, Address:="", SubAddress:= _
 ws.Name & "!A1", TextToDisplay:=ws.Name
 x = x + 1
 
Next ws
 
End Sub

'add to worksheet where button is located

Private Sub btnRefreshList_Click()
    
    Call ListSheets
    ActiveSheet.Range("B7").Select
    
End Sub
