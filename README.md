# April30_04_2022
Excel

HiSub Check_Url()
    Dim i%, row%, column%, import, output
    import = ""
    output = ""
    With ActiveSheet.UsedRange
        row = .Rows.Count + .row - 1
        column = .Columns.Count + .column - 1
        If import = "" Then import = "a"
        If output = "" Then output = column
    End With
    For i = 2 To row
        With CreateObject("Microsoft.XMLHTTP")
            .Open "get", Cells(i, import), False
            .send
            If .Status = 404 Then
                Cells(i, output) = "无效"
            Else
                Cells(i, output) = "有效"
            End If
        End With
    Next i
End Sub




https://github.com/bgoonz/WEB-DEV-NOTES/blob/47df2e93fee49054074dc0cffba970f5d6096124/00-4-all-time/general-ref/VBA-master/VBA-master/All%20BAS%20Files/HTML_VBA_Excel.bas
