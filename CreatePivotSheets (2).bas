Attribute VB_Name = "CreatePivotSheets"
Option Explicit

Sub CreatePivotSheets()

    Dim wsData As Worksheet
    Dim wsPvt As Worksheet
    Dim pc As PivotCache
    Dim lastRow As Long
    Dim dataRange As Range

    Dim ws As Worksheet
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("Data")
    On Error GoTo 0

    If wsData Is Nothing Then
        MsgBox "Sheet named Data not found.", vbCritical, "Error"
        Exit Sub
    End If

    For Each ws In ThisWorkbook.Worksheets
        If LCase(Trim(ws.Name)) = "pivot" Then
            Set wsPvt = ws
            Exit For
        End If
    Next ws

    If wsPvt Is Nothing Then
        MsgBox "Sheet named Pivot not found.", vbCritical, "Error"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Do While wsPvt.PivotTables.Count > 0
        wsPvt.PivotTables(1).TableRange2.Clear
    Loop
    wsPvt.Cells.Clear

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set dataRange = wsData.Range("A1:M" & lastRow)
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)

    Call BuildPT_Segment(pc, wsPvt)
    Call BuildPT_Trend(pc, wsPvt)
    Call BuildPT_Monthly(pc, wsPvt)
    Call BuildPT_Category(pc, wsPvt)
    Call BuildPT_Region(pc, wsPvt)
    Call BuildPT_Discount(pc, wsPvt)
    Call BuildPT_Top5(pc, wsPvt)
    Call BuildPT_Quantity(pc, wsPvt)
    Call WritePivotTitles(wsPvt)

    wsPvt.Columns("A:AE").AutoFit
    wsPvt.Activate
    ActiveWindow.FreezePanes = False
    wsPvt.Range("A3").Select
    ActiveWindow.FreezePanes = True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Pivot sheet built successfully!", vbInformation, "Done"

End Sub


Private Sub BuildPT_Segment(pc As PivotCache, wsPvt As Worksheet)
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsPvt.Range("A2"), TableName:="PT_Segment")
    With pt
        .PivotFields("Segment").Orientation = xlRowField
        .PivotFields("Segment").Position = 1
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Total Sales"
            .NumberFormat = "#,##0.00"
        End With
    End With
End Sub


Private Sub BuildPT_Trend(pc As PivotCache, wsPvt As Worksheet)
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsPvt.Range("D2"), TableName:="PT_Trend")
    With pt
        .PivotFields("Year").Orientation = xlRowField
        .PivotFields("Year").Position = 1
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Month").Position = 2
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sum of Sales"
            .NumberFormat = "#,##0.00"
        End With
        With .PivotFields("Profit")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sum of Profit"
            .NumberFormat = "#,##0.00"
        End With
    End With
End Sub


Private Sub BuildPT_Monthly(pc As PivotCache, wsPvt As Worksheet)
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsPvt.Range("H2"), TableName:="PT_Monthly")
    With pt
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Month").Position = 1
        .PivotFields("Year").Orientation = xlColumnField
        .PivotFields("Year").Position = 1
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sum of Sales"
            .NumberFormat = "#,##0.00"
        End With
    End With
End Sub


Private Sub BuildPT_Category(pc As PivotCache, wsPvt As Worksheet)
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsPvt.Range("O2"), TableName:="PT_Category")
    With pt
        .PivotFields("Category").Orientation = xlRowField
        .PivotFields("Category").Position = 1
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sum of Sales"
            .NumberFormat = "#,##0.00"
        End With
    End With
End Sub


Private Sub BuildPT_Region(pc As PivotCache, wsPvt As Worksheet)
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsPvt.Range("R2"), TableName:="PT_Region")
    With pt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Region").Position = 1
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sum of Sales"
            .NumberFormat = "#,##0.00"
        End With
    End With
End Sub


Private Sub BuildPT_Discount(pc As PivotCache, wsPvt As Worksheet)
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsPvt.Range("U2"), TableName:="PT_Discount")
    With pt
        .PivotFields("Sub-Category").Orientation = xlRowField
        .PivotFields("Sub-Category").Position = 1
        With .PivotFields("Profit")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sum of Profit"
            .NumberFormat = "#,##0.00"
        End With
        With .PivotFields("Discount")
            .Orientation = xlDataField
            .Function = xlAverage
            .Name = "Avg Discount"
            .NumberFormat = "0.00%"
        End With
    End With
End Sub


Private Sub BuildPT_Top5(pc As PivotCache, wsPvt As Worksheet)
    Dim pt As PivotTable
    Dim fName As String
    Set pt = pc.CreatePivotTable(TableDestination:=wsPvt.Range("Y2"), TableName:="PT_Top5")
    With pt
        .PivotFields("Order ID").Orientation = xlRowField
        .PivotFields("Order ID").Position = 1
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sum of Sales"
            .NumberFormat = "#,##0.00"
        End With
        fName = pt.DataFields(1).Name
        .PivotFields("Order ID").AutoSort xlDescending, fName
    End With
End Sub


Private Sub BuildPT_Quantity(pc As PivotCache, wsPvt As Worksheet)
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsPvt.Range("AB2"), TableName:="PT_Quantity")
    With pt
        .PivotFields("Segment").Orientation = xlRowField
        .PivotFields("Segment").Position = 1
        With .PivotFields("Quantity")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Total Quantity"
            .NumberFormat = "#,##0.00"
        End With
    End With
End Sub


Sub WritePivotTitles(wsPvt As Worksheet)

    Dim purple As Long
    Dim bgColor As Long
    Dim white As Long
    purple = RGB(94, 59, 183)
    bgColor = RGB(235, 228, 255)
    white = RGB(255, 255, 255)

    Dim titles(7, 2) As String
    titles(0, 0) = "Sales by Segment":         titles(0, 1) = "A":  titles(0, 2) = "B"
    titles(1, 0) = "Sales and Profit Trend":   titles(1, 1) = "D":  titles(1, 2) = "F"
    titles(2, 0) = "Monthly Sales":            titles(2, 1) = "H":  titles(2, 2) = "M"
    titles(3, 0) = "Sales by Category":        titles(3, 1) = "O":  titles(3, 2) = "P"
    titles(4, 0) = "Sales by Region":          titles(4, 1) = "R":  titles(4, 2) = "S"
    titles(5, 0) = "Profit vs Discount":       titles(5, 1) = "U":  titles(5, 2) = "W"
    titles(6, 0) = "Top 5 Sales Orders":       titles(6, 1) = "Y":  titles(6, 2) = "Z"
    titles(7, 0) = "Sales by Quantity":        titles(7, 1) = "AB": titles(7, 2) = "AC"

    Dim i As Integer
    Dim mergeRng As Range
    Dim borderRng As Range

    For i = 0 To 7
        Set mergeRng = wsPvt.Range(titles(i, 1) & "1:" & titles(i, 2) & "1")
        On Error Resume Next
        mergeRng.UnMerge
        On Error GoTo 0
        mergeRng.Merge
        With mergeRng
            .Value = titles(i, 0)
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = white
            .Font.Name = "Calibri"
            .Interior.Color = purple
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Set borderRng = wsPvt.Range( _
            wsPvt.Cells(2, mergeRng.Column), _
            wsPvt.Cells(2, mergeRng.Column + mergeRng.Columns.Count - 1))
        With borderRng.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = bgColor
        End With
    Next i

    wsPvt.Rows(1).RowHeight = 20

End Sub


Sub RefreshAllPivots()
    Dim pt As PivotTable
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    MsgBox "All Pivot Tables refreshed!", vbInformation, "Done"
End Sub
