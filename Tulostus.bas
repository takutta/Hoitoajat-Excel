Attribute VB_Name = "Tulostus"
Sub Tulosta()

    Dim WSname As Variant
    For Each WSname In Array("Code")
        Worksheets(WSname).Visible = True
    Next

    Dim wb As Workbook: Set wb = ThisWorkbook

    Dim sht_code As Worksheet: Set sht_code = wb.Worksheets("Code")

    Dim lRow As Long: lRow = sht_code.Cells(Rows.Count, 2).End(xlUp).Row
    Dim vanharyh As Range

    On Error Resume Next

    Dim i As Long
    If lRow <> 1 Then
        sht_code.Select
        For i = sht_code.Range(Cells(2, 2), Cells(lRow, 2)).Count To 1 Step -1
            wb.Worksheets(sht_code.Cells(1 + i, 2).Value).PrintOut Copies:=1, Preview:=False, Collate:=True
            wb.Worksheets(sht_code.Cells(1 + i, 2).Value & "_ma").PrintOut Copies:=1, Preview:=False, Collate:=True
            wb.Worksheets(sht_code.Cells(1 + i, 2).Value & "_ti").PrintOut Copies:=1, Preview:=False, Collate:=True
            wb.Worksheets(sht_code.Cells(1 + i, 2).Value & "_ke").PrintOut Copies:=1, Preview:=False, Collate:=True
            wb.Worksheets(sht_code.Cells(1 + i, 2).Value & "_to").PrintOut Copies:=1, Preview:=False, Collate:=True
            wb.Worksheets(sht_code.Cells(1 + i, 2).Value & "_pe").PrintOut Copies:=1, Preview:=False, Collate:=True
        Next i
        wb.Worksheets("Aamulista").PrintOut Copies:=1, Preview:=False, Collate:=True
        wb.Worksheets("Iltalista").PrintOut Copies:=1, Preview:=False, Collate:=True
    End If

    wb.Worksheets("Päiväkoti").Select

    Sheets(Array("Code")).Select
    ActiveWindow.SelectedSheets.Visible = False

End Sub

