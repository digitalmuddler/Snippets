Sub MakeRange() 
 
    ActiveWorkbook.Names.Add _ 
        Name:="tempRange", _ 
        RefersTo:="=Sheet1!$A$1:$D$3" 
 
End Sub
