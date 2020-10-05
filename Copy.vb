Sub copy()
Dim Wb As Workbook
Dim Wb2 As Workbook
Set Wb = ThisWorkbook

    ' Range("A2:B2" & Range("A" & Rows.Count).End(xlUp).Row).FillDown
        Workbooks.Open ("\\SERVER2019\Data\Finance\Carriage\for macro\Checking file.xlsx")
        Set Wb2 = ThisWorkbook
        Wb.Activate
        Wb.Worksheets("PVX Data").Range("A2:BB10000").copy
            Wb2.Activate
            
            Wb2.Worksheets("Despatch summary").Range("A2:BB10000").PasteSpecial Paste:=xlPasteValues

End Sub
