Attribute VB_Name = "Module1"
Sub filtering_ISO()
Attribute filtering_ISO.VB_ProcData.VB_Invoke_Func = " \n14"
With Worksheets("Sheet1")
    .AutoFilterMode = Fals
    .Range("B2").AutoFilter Field:=1, Criteria1:="=*ISO*"
End With
End Sub
Sub filtering_IEC()
With Worksheets("Sheet1")
    .AutoFilterMode = Fals
    .Range("B2").AutoFilter Field:=1, Criteria1:="=*IEC*"
End With
    Range("B2").Select
End Sub
Sub release_filtering()
With Worksheets("Sheet1")
    If .AutoFilterMode Then
        .AutoFilterMode = False
    End If
End With
    Range("B2").Select
End Sub
