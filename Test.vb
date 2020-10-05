Sub Test  
Dim wkb as Workbook
Dim wkbFrom as Worksheet
set wkb = ThisWorkbook
   
    ActiveCell.Select
    ActiveCell.Formula2R1C1 = "hello world I am testing VB"
    With Selction.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
With Selction.Font
    .color = 4165632
    .TintAndShade = 0
End With

End Sub
Sub Fill_Down()
  Range("A2:B2" & Range("A" & Rows.Count).End(xlUp).Row).FillDown
End Sub

End Sub
Sub clearcontent()
  With sheets("Total").Select
    If MsgBox("This will clear all your contents, are you SURE you want to do this?", vbYesNo, "Confirm") = vbYes Then
    Range("C2:BV2" & Range("A" & Rows.Count).End(xlUp).Row).ClearContents
    End If
  End With
End Sub

Sub Test()
  With ActiveSheet
    MsgBox "Lines are updated"
      If 1 = 1 Then
        If MsgBox("put something in the sheet", vbYesNo, "Confirm") = vbYes Then
        Loop
      End If
  End With