Attribute VB_Name = "CleanAndFormatData"
Option Explicit

Sub CleanAndFormatData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim recordCount As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Data")
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        MsgBox "Sheet named Data not found.", vbCritical, "Error"
        GoTo CleanUp
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Call StepRemoveBlankRowsCols(ws, lastRow, lastCol)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Call StepTrimWhitespace(ws, lastRow, lastCol)
    Call StepStandardiseText(ws, lastRow)
    Call StepFormatNumbers(ws, lastRow)
    Call StepCentreAlignment(ws, lastRow, lastCol)
    Call StepApplyPurpleTheme(ws, lastRow, lastCol)
    Call StepFormatHeader(ws, lastCol)
    Call StepSetColumnWidths(ws)
    Call StepFreezeHeader(ws)

    recordCount = lastRow - 1
    MsgBox "Done! Records processed: " & recordCount, vbInformation, "Complete"

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    Resume CleanUp

End Sub


Sub StepRemoveBlankRowsCols(ws As Worksheet, lastRow As Long, lastCol As Long)

    Dim r As Long
    Dim c As Long

    For r = lastRow To 2 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
        End If
    Next r

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = lastCol To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Columns(c)) = 0 Then
            ws.Columns(c).Delete
        End If
    Next c

End Sub


Sub StepTrimWhitespace(ws As Worksheet, lastRow As Long, lastCol As Long)

    Dim r As Long
    Dim c As Long
    Dim cel As Range

    For r = 1 To lastRow
        For c = 1 To lastCol
            Set cel = ws.Cells(r, c)
            If VarType(cel.Value) = vbString Then
                If cel.Value <> "" Then
                    cel.Value = Trim(cel.Value)
                End If
            End If
        Next c
    Next r

End Sub


Sub StepStandardiseText(ws As Worksheet, lastRow As Long)

    Dim colIdx As Integer
    Dim r As Long
    Dim headers As Variant
    Dim h As Variant

    headers = Array("Segment", "Category", "Sub-Category", "Region")

    For Each h In headers
        colIdx = FindColumn(ws, CStr(h))
        If colIdx > 0 Then
            For r = 2 To lastRow
                If ws.Cells(r, colIdx).Value <> "" Then
                    ws.Cells(r, colIdx).Value = Application.WorksheetFunction.Proper(ws.Cells(r, colIdx).Value)
                End If
            Next r
        End If
    Next h

End Sub


Sub StepFormatNumbers(ws As Worksheet, lastRow As Long)

    Dim r As Long
    Dim colIdx As Integer
    Dim cel As Range

    colIdx = FindColumn(ws, "Sales")
    If colIdx > 0 Then
        ws.Range(ws.Cells(2, colIdx), ws.Cells(lastRow, colIdx)).NumberFormat = "#,##0.00"
    End If

    colIdx = FindColumn(ws, "Profit")
    If colIdx > 0 Then
        ws.Range(ws.Cells(2, colIdx), ws.Cells(lastRow, colIdx)).NumberFormat = "#,##0.00"
    End If

    colIdx = FindColumn(ws, "Quantity")
    If colIdx > 0 Then
        ws.Range(ws.Cells(2, colIdx), ws.Cells(lastRow, colIdx)).NumberFormat = "#,##0"
    End If

    colIdx = FindColumn(ws, "Discount")
    If colIdx > 0 Then
        ws.Range(ws.Cells(2, colIdx), ws.Cells(lastRow, colIdx)).NumberFormat = "0%"
    End If

    colIdx = FindColumn(ws, "Year")
    If colIdx > 0 Then
        ws.Range(ws.Cells(2, colIdx), ws.Cells(lastRow, colIdx)).NumberFormat = "0"
    End If

    colIdx = FindColumn(ws, "Day")
    If colIdx > 0 Then
        ws.Range(ws.Cells(2, colIdx), ws.Cells(lastRow, colIdx)).NumberFormat = "0"
    End If

    colIdx = FindColumn(ws, "Order Date")
    If colIdx > 0 Then
        For r = 2 To lastRow
            Set cel = ws.Cells(r, colIdx)
            If IsNumeric(cel.Value) Then
                If CDbl(cel.Value) > 1 Then
                    cel.Value = CDate(cel.Value)
                End If
            End If
            cel.NumberFormat = "yyyy-mm-dd"
        Next r
    End If

End Sub


Sub StepCentreAlignment(ws As Worksheet, lastRow As Long, lastCol As Long)

    Dim fullRange As Range
    Set fullRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    With fullRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 18
    End With

End Sub


Sub StepApplyPurpleTheme(ws As Worksheet, lastRow As Long, lastCol As Long)

    Dim dataRange As Range
    Dim lo As ListObject
    Dim r As Long

    Set dataRange = ws.Range("A1").CurrentRegion

    On Error Resume Next
    For Each lo In ws.ListObjects
        lo.Unlist
    Next lo
    On Error GoTo 0

    Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)

    lo.Name = "tblSalesData"
    lo.TableStyle = "TableStyleMedium13"
    lo.ShowAutoFilter = True

    For r = 2 To lastRow
        If r Mod 2 <> 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(235, 228, 255)
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(255, 255, 255)
        End If
    Next r

    dataRange.Columns.AutoFit

End Sub


Sub StepFormatHeader(ws As Worksheet, lastCol As Long)

    Dim headerRange As Range
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))

    With headerRange
        .Interior.Color = RGB(94, 59, 183)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With

End Sub


Sub StepSetColumnWidths(ws As Worksheet)

    Dim i As Integer
    Dim colIdx As Integer
    Dim names(12) As String
    Dim widths(12) As Double

    names(0) = "Order ID"
    names(1) = "Order Date"
    names(2) = "Year"
    names(3) = "Month"
    names(4) = "Day"
    names(5) = "Segment"
    names(6) = "Category"
    names(7) = "Sub-Category"
    names(8) = "Region"
    names(9) = "Sales"
    names(10) = "Profit"
    names(11) = "Quantity"
    names(12) = "Discount"

    widths(0) = 13
    widths(1) = 14
    widths(2) = 8
    widths(3) = 9
    widths(4) = 7
    widths(5) = 14
    widths(6) = 16
    widths(7) = 16
    widths(8) = 10
    widths(9) = 12
    widths(10) = 12
    widths(11) = 10
    widths(12) = 10

    For i = 0 To 12
        colIdx = FindColumn(ws, names(i))
        If colIdx > 0 Then
            ws.Columns(colIdx).ColumnWidth = widths(i)
        End If
    Next i

End Sub


Sub StepFreezeHeader(ws As Worksheet)

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

End Sub


Function FindColumn(ws As Worksheet, headerName As String) As Integer

    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If LCase(Trim(ws.Cells(1, c).Value)) = LCase(Trim(headerName)) Then
            FindColumn = CInt(c)
            Exit Function
        End If
    Next c

    FindColumn = 0

End Function
