Attribute VB_Name = "DashboardModule"
Option Explicit

' ============================================================
'  BuildDashboard v7  -  Complete Rewrite
'  KPI cards fully fixed: no gaps, centred title, clean layout
' ============================================================

Public Sub BuildDashboard()

    Dim wb     As Workbook
    Dim wsDash As Worksheet
    Set wb = ThisWorkbook

    ' ── Verify Pivot sheet exists ────────────────────────────
    Dim pivotFound As Boolean : pivotFound = False
    Dim wsChk As Worksheet
    For Each wsChk In wb.Worksheets
        If wsChk.Name = "Pivot" Then pivotFound = True : Exit For
    Next wsChk
    If Not pivotFound Then
        MsgBox "ERROR: Sheet named 'Pivot' not found!" & Chr(10) & _
               "Rename your pivot sheet to: Pivot", vbCritical, "Missing Sheet"
        Exit Sub
    End If

    ' ── Delete and recreate Dashboard sheet (avoids merge conflicts) ──
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = "Dashboard" Then ws.Delete : Exit For
    Next ws
    Application.DisplayAlerts = True
    Set wsDash = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsDash.Name = "Dashboard"

    ' ── Column widths: 24 cols (A:X), 4 cols per KPI card ───
    Dim c As Long
    For c = 1 To 24
        wsDash.Columns(c).ColumnWidth = 11
    Next c

    ' ── Row heights ──────────────────────────────────────────
    wsDash.Rows(1).RowHeight  = 28   ' Header top
    wsDash.Rows(2).RowHeight  = 28   ' Header bottom
    wsDash.Rows(3).RowHeight  = 4    ' Accent stripe
    wsDash.Rows(4).RowHeight  = 20   ' Card icon row
    wsDash.Rows(5).RowHeight  = 36   ' Card value row (big number)
    wsDash.Rows(6).RowHeight  = 18   ' Card label row
    wsDash.Rows(7).RowHeight  = 5    ' Card bottom bar
    wsDash.Rows(8).RowHeight  = 16   ' Section label
    Dim r As Long
    For r = 9  To 24 : wsDash.Rows(r).RowHeight = 15 : Next r   ' Top charts
    wsDash.Rows(25).RowHeight = 6                                ' Spacer
    For r = 26 To 41 : wsDash.Rows(r).RowHeight = 15 : Next r   ' Bottom charts
    wsDash.Rows(42).RowHeight = 16                               ' Footer

    ' ── Full background ──────────────────────────────────────
    wsDash.Range("A1:X42").Interior.Color = RGB(18, 18, 35)

    ' ====================================================
    '  HEADER  (rows 1-2, centred title)
    ' ====================================================
    wsDash.Range("A1:X2").Merge
    With wsDash.Range("A1")
        .Value = "SALES PERFORMANCE DASHBOARD"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
        .Interior.Color      = RGB(94, 59, 183)
        .Font.Name  = "Segoe UI"
        .Font.Size  = 26
        .Font.Bold  = True
        .Font.Color = RGB(255, 255, 255)
    End With
    ' Accent stripe row 3
    wsDash.Range("A3:X3").Interior.Color = RGB(138, 99, 228)

    ' ====================================================
    '  KPI CARDS  (rows 4-7, 6 cards x 4 cols = 24 cols)
    ' ====================================================
    '  Structure of each card (4 cols wide):
    '    Col cs   = left accent bar (coloured)
    '    Col cs+1 = icon (row4), value (row5), label (row6)
    '    Col cs+2 = value overflow / empty
    '    Col cs+3 = right padding
    '    Row 7    = bottom accent bar (full width)

    Dim kStart(5) As Long   ' starting column of each card
    kStart(0) = 1   ' A
    kStart(1) = 5   ' E
    kStart(2) = 9   ' I
    kStart(3) = 13  ' M
    kStart(4) = 17  ' Q
    kStart(5) = 21  ' U

    Dim kLabel(5) As String
    kLabel(0) = "TOTAL SALES"
    kLabel(1) = "TOTAL PROFIT"
    kLabel(2) = "TOTAL ORDERS"
    kLabel(3) = "TOTAL QTY SOLD"
    kLabel(4) = "AVG DISCOUNT"
    kLabel(5) = "PROFIT MARGIN"

    ' kStart(0) val cell = col 2 = B, row5
    ' kStart(1) val cell = col 6 = F, row5
    Dim kFormula(5) As String
    kFormula(0) = "=IFERROR(GETPIVOTDATA(""Sum of Sales"",Pivot!$A$3),IFERROR(GETPIVOTDATA(""Total Sales"",Pivot!$A$3),0))"
    kFormula(1) = "=IFERROR(GETPIVOTDATA(""Sum of Profit"",Pivot!$D$3),IFERROR(GETPIVOTDATA(""Total Profit"",Pivot!$D$3),0))"
    kFormula(2) = "=IFERROR(COUNTA(Orders!A:A)-1,IFERROR(COUNTA(Data!A:A)-1,IFERROR(COUNTA(Sheet1!A:A)-1,0)))"
    kFormula(3) = "=IFERROR(GETPIVOTDATA(""Total Quantity"",Pivot!$AB$3),IFERROR(GETPIVOTDATA(""Sum of Quantity"",Pivot!$AB$3),0))"
    kFormula(4) = "=IFERROR(GETPIVOTDATA(""Avg Discount"",Pivot!$U$3),IFERROR(GETPIVOTDATA(""Average of Discount"",Pivot!$U$3),0))"
    kFormula(5) = "=IFERROR(F5/B5,0)"

    Dim kFmt(5) As String
    kFmt(0) = "#,##0" : kFmt(1) = "#,##0" : kFmt(2) = "#,##0"
    kFmt(3) = "#,##0" : kFmt(4) = "0.0%"  : kFmt(5) = "0.0%"

    Dim kIcon(5) As String
    kIcon(0) = Chr(36)      ' $
    kIcon(1) = ChrW(8593)   ' up arrow
    kIcon(2) = ChrW(10003)  ' checkmark
    kIcon(3) = ChrW(9632)   ' filled square
    kIcon(4) = Chr(37)      ' %
    kIcon(5) = ChrW(9711)   ' circle

    Dim kColor(5) As Long
    kColor(0) = RGB(94,  59, 183)   ' purple
    kColor(1) = RGB(0,  200, 150)   ' teal
    kColor(2) = RGB(255, 140,   0)  ' orange
    kColor(3) = RGB(68,  114, 196)  ' blue
    kColor(4) = RGB(197,  90, 240)  ' violet
    kColor(5) = RGB(255,  80, 120)  ' coral

    Dim i As Long, cs As Long
    For i = 0 To 5
        cs = kStart(i)

        ' --- Full card background rows 4-7, cols cs to cs+3 ---
        With wsDash.Range(wsDash.Cells(4, cs), wsDash.Cells(7, cs + 3))
            .Interior.Color = RGB(28, 25, 50)
        End With

        ' --- Left accent bar (col cs, rows 4-6) ---
        With wsDash.Range(wsDash.Cells(4, cs), wsDash.Cells(6, cs))
            .Interior.Color = kColor(i)
        End With

        ' --- Row 4: Icon ---
        With wsDash.Cells(4, cs + 1)
            .Value              = kIcon(i)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment  = xlBottom
            .Font.Color         = kColor(i)
            .Font.Size          = 13
            .Font.Bold          = True
            .IndentLevel        = 1
        End With

        ' --- Row 5: Value (merge cs+1 : cs+3 for full width) ---
        wsDash.Range(wsDash.Cells(5, cs + 1), wsDash.Cells(5, cs + 3)).Merge
        With wsDash.Cells(5, cs + 1)
            .Formula            = kFormula(i)
            .NumberFormat       = kFmt(i)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment  = xlCenter
            .Font.Name          = "Segoe UI"
            .Font.Size          = 20
            .Font.Bold          = True
            .Font.Color         = RGB(255, 255, 255)
            .IndentLevel        = 1
        End With

        ' --- Row 6: Label ---
        With wsDash.Cells(6, cs + 1)
            .Value              = kLabel(i)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment  = xlTop
            .Font.Name          = "Segoe UI"
            .Font.Size          = 8
            .Font.Bold          = True
            .Font.Color         = RGB(160, 155, 200)
            .IndentLevel        = 1
        End With

        ' --- Row 7: Full-width bottom accent bar ---
        With wsDash.Range(wsDash.Cells(7, cs), wsDash.Cells(7, cs + 3))
            .Interior.Color = kColor(i)
        End With

        ' --- Outer border around whole card ---
        With wsDash.Range(wsDash.Cells(4, cs), wsDash.Cells(7, cs + 3)).Borders
            .LineStyle = xlContinuous
            .Weight    = xlThin
            .Color     = kColor(i)
        End With
    Next i

    ' ====================================================
    '  SECTION LABEL  (row 8)
    ' ====================================================
    With wsDash.Cells(8, 2)
        .Value = "  PERFORMANCE ANALYTICS"
        .Font.Name  = "Segoe UI"
        .Font.Size  = 8
        .Font.Bold  = True
        .Font.Color = RGB(138, 99, 228)
        .VerticalAlignment = xlCenter
    End With

    ' ====================================================
    '  FOOTER  (row 42)
    ' ====================================================
    wsDash.Range("A42:X42").Interior.Color = RGB(94, 59, 183)
    With wsDash.Cells(42, 2)
        .Value = "  Sales Performance Dashboard  |  Powered by Excel VBA  |  Data Source: Pivot Sheet"
        .Font.Name  = "Segoe UI"
        .Font.Size  = 8
        .Font.Color = RGB(220, 215, 255)
        .VerticalAlignment = xlCenter
    End With

    ' ====================================================
    '  CHARTS
    ' ====================================================
    Call CreateAllCharts(wsDash)

    wsDash.Activate
    wsDash.Tab.Color = RGB(94, 59, 183)
    ActiveWindow.DisplayGridlines = False
    wsDash.Range("A1").Select
    MsgBox "Dashboard v7 built!", vbInformation, "Done"

End Sub


' ============================================================
Private Sub ApplyTitleStyle(ct As ChartTitle, titleText As String)
    ct.Text = titleText
    With ct.Font
        .Name = "Segoe UI" : .Size = 10 : .Bold = True
        .Color = RGB(220, 215, 255)
    End With
    ct.Left = 8
End Sub

Private Sub ApplyChartAreaStyle(cht As Chart)
    With cht.ChartArea
        .Interior.Color = RGB(22, 20, 42)
        With .Border
            .LineStyle = xlContinuous : .Weight = xlThin : .Color = RGB(94, 59, 183)
        End With
        With .Font
            .Name = "Segoe UI" : .Size = 8 : .Color = RGB(180, 175, 220)
        End With
    End With
    With cht.PlotArea
        .Interior.Color = RGB(30, 28, 52)
        .Border.LineStyle = xlNone
    End With
End Sub

Private Sub StyleAxes(cht As Chart)
    On Error Resume Next
    Dim ax As Axis
    For Each ax In cht.Axes
        ax.MajorGridlines.Delete : ax.MinorGridlines.Delete
        With ax.TickLabels.Font
            .Color = RGB(160, 155, 200) : .Size = 7 : .Name = "Segoe UI"
        End With
        ax.AxisTitle.Delete
        With ax.Border
            .LineStyle = xlContinuous : .Color = RGB(60, 55, 90) : .Weight = xlHairline
        End With
    Next ax
    On Error GoTo 0
End Sub

Private Sub AddValueGridlines(cht As Chart)
    On Error Resume Next
    With cht.Axes(xlValue).MajorGridlines
        .Format.Line.ForeColor.RGB = RGB(50, 45, 80)
        .Format.Line.Weight = 0.5
        .Format.Line.DashStyle = msoLineDash
    End With
    On Error GoTo 0
End Sub

Private Sub HideFieldButtons(cht As Chart)
    On Error Resume Next
    cht.ShowAllFieldButtons = False
    cht.ShowAxisFieldButtons = False
    cht.ShowLegendFieldButtons = False
    On Error GoTo 0
End Sub

Private Sub StyleLegend(cht As Chart, pos As Long)
    On Error Resume Next
    cht.HasLegend = True
    With cht.Legend
        .Position = pos
        With .Font
            .Color = RGB(160, 155, 200) : .Size = 7 : .Name = "Segoe UI"
        End With
        .Border.LineStyle = xlNone
        .Interior.ColorIndex = xlNone
    End With
    On Error GoTo 0
End Sub


' ============================================================
Private Sub CreateAllCharts(wsDash As Worksheet)

    Dim cPurple As Long : cPurple = RGB(94,  59, 183)
    Dim cViolet As Long : cViolet = RGB(138,  99, 228)
    Dim cPink   As Long : cPink   = RGB(197,  90, 240)
    Dim cOrange As Long : cOrange = RGB(255, 140,   0)
    Dim cBlue   As Long : cBlue   = RGB(68,  114, 196)
    Dim cTeal   As Long : cTeal   = RGB(0,   200, 150)
    Dim cCoral  As Long : cCoral  = RGB(255,  80, 120)
    Dim cWhite  As Long : cWhite  = RGB(220, 215, 255)

    Dim wb      As Workbook  : Set wb      = wsDash.Parent
    Dim wsPivot As Worksheet : Set wsPivot = wb.Worksheets("Pivot")

    Const GAP As Double = 4

    ' 3 chart columns across 24 cols: cols 1-8, 9-16, 17-24
    Dim c1L As Double : c1L = wsDash.Cells(1,  1).Left + GAP
    Dim c1R As Double : c1R = wsDash.Cells(1,  9).Left - GAP
    Dim c2L As Double : c2L = wsDash.Cells(1,  9).Left + GAP
    Dim c2R As Double : c2R = wsDash.Cells(1, 17).Left - GAP
    Dim c3L As Double : c3L = wsDash.Cells(1, 17).Left + GAP
    Dim c3R As Double : c3R = wsDash.Cells(1, 25).Left - GAP

    ' Top charts: rows 9-24 | Bottom charts: rows 26-41
    Dim tT As Double : tT = wsDash.Cells(9,  1).Top + GAP
    Dim tB As Double : tB = wsDash.Cells(25, 1).Top - GAP
    Dim bT As Double : bT = wsDash.Cells(26, 1).Top + GAP
    Dim bB As Double : bB = wsDash.Cells(42, 1).Top - GAP

    Dim w1 As Double : w1 = c1R - c1L
    Dim w2 As Double : w2 = c2R - c2L
    Dim w3 As Double : w3 = c3R - c3L
    Dim hT As Double : hT = tB - tT
    Dim hB As Double : hB = bB - bT

    Dim cht As Chart
    Dim pt  As Long
    Dim s   As Long

    ' CHART 1 - Sales by Segment (Doughnut) - Top Left
    With wsDash.ChartObjects.Add(c1L, tT, w1, hT) : Set cht = .Chart : End With
    With cht
        .SetSourceData Source:=wsPivot.Range("N3:O6"), PlotBy:=xlColumns
        .ChartType = xlDoughnut : .HasTitle = True
        ApplyTitleStyle .ChartTitle, "Sales by Segment"
        ApplyChartAreaStyle cht : StyleAxes cht : HideFieldButtons cht
        StyleLegend cht, xlLegendPositionBottom
        On Error Resume Next
        .ChartGroups(1).DoughnutHoleSize = 60
        Dim segC(2) As Long
        segC(0) = cPurple : segC(1) = cOrange : segC(2) = cPink
        For pt = 1 To 3
            .SeriesCollection(1).Points(pt).Interior.Color = segC(pt - 1)
            .SeriesCollection(1).Points(pt).Border.LineStyle = xlNone
        Next pt
        .SeriesCollection(1).HasDataLabels = True
        With .SeriesCollection(1).DataLabels
            .ShowPercentage = True : .ShowValue = False : .ShowCategoryName = True
            .Separator = Chr(10) : .Position = xlLabelPositionBestFit
            With .Font : .Color = cWhite : .Bold = True : .Size = 8 : .Name = "Segoe UI" : End With
        End With
        On Error GoTo 0
    End With

    ' CHART 2 - Sales & Profit Trend (Line) - Top Middle
    With wsDash.ChartObjects.Add(c2L, tT, w2, hT) : Set cht = .Chart : End With
    With cht
        .SetSourceData Source:=wsPivot.Range("D3:F55"), PlotBy:=xlColumns
        .ChartType = xlLine : .HasTitle = True
        ApplyTitleStyle .ChartTitle, "Sales & Profit Trend"
        ApplyChartAreaStyle cht : StyleAxes cht : AddValueGridlines cht
        HideFieldButtons cht : StyleLegend cht, xlLegendPositionBottom
        On Error Resume Next
        With .SeriesCollection(1)
            .Format.Line.ForeColor.RGB = cViolet : .Format.Line.Weight = 2
            .Format.Line.Transparency = 0 : .MarkerStyle = xlMarkerStyleNone : .Name = "Sales"
        End With
        With .SeriesCollection(2)
            .Format.Line.ForeColor.RGB = cTeal : .Format.Line.Weight = 1.5
            .Format.Line.DashStyle = msoLineDash : .MarkerStyle = xlMarkerStyleNone : .Name = "Profit"
        End With
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        On Error GoTo 0
    End With

    ' CHART 3 - Monthly Sales (Clustered Column) - Top Right
    With wsDash.ChartObjects.Add(c3L, tT, w3, hT) : Set cht = .Chart : End With
    With cht
        .SetSourceData Source:=wsPivot.Range("H3:L15"), PlotBy:=xlColumns
        .ChartType = xlColumnClustered : .HasTitle = True
        ApplyTitleStyle .ChartTitle, "Monthly Sales Performance"
        ApplyChartAreaStyle cht : StyleAxes cht : AddValueGridlines cht
        HideFieldButtons cht : StyleLegend cht, xlLegendPositionBottom
        On Error Resume Next
        Dim mC(3) As Long
        mC(0) = cPurple : mC(1) = cViolet : mC(2) = cOrange : mC(3) = cBlue
        For s = 1 To .SeriesCollection.Count
            If s <= 4 Then
                .SeriesCollection(s).Format.Fill.ForeColor.RGB = mC(s - 1)
                .SeriesCollection(s).Format.Fill.Transparency = 0
                .SeriesCollection(s).Interior.Color = mC(s - 1)
            End If
        Next s
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .ChartGroups(1).GapWidth = 80
        On Error GoTo 0
    End With

    ' CHART 4 - Profit by Sub-Category (Bar) - Bottom Left
    With wsDash.ChartObjects.Add(c1L, bT, w1, hB) : Set cht = .Chart : End With
    With cht
        .SetSourceData Source:=wsPivot.Range("U3:V20"), PlotBy:=xlColumns
        .ChartType = xlBarClustered : .HasTitle = True
        ApplyTitleStyle .ChartTitle, "Profit by Sub-Category"
        ApplyChartAreaStyle cht : StyleAxes cht : AddValueGridlines cht
        HideFieldButtons cht : cht.HasLegend = False
        On Error Resume Next
        With .SeriesCollection(1)
            .Interior.Color = cViolet
            .Format.Fill.ForeColor.RGB = cViolet
            .Format.Fill.Transparency = 0.15
        End With
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        With .Axes(xlCategory).TickLabels.Font
            .Size = 7 : .Color = RGB(180, 175, 220) : .Name = "Segoe UI"
        End With
        .ChartGroups(1).GapWidth = 25
        With cht.PlotArea : .InsideLeft = 85 : End With
        On Error GoTo 0
    End With

    ' CHART 5 - Sales by Category (Bar) - Bottom Middle
    With wsDash.ChartObjects.Add(c2L, bT, w2, hB) : Set cht = .Chart : End With
    With cht
        .SetSourceData Source:=wsPivot.Range("O3:P6"), PlotBy:=xlColumns
        .ChartType = xlBarClustered : .HasTitle = True
        ApplyTitleStyle .ChartTitle, "Sales by Category"
        ApplyChartAreaStyle cht : StyleAxes cht : HideFieldButtons cht : cht.HasLegend = False
        On Error Resume Next
        Dim catC(2) As Long
        catC(0) = cTeal : catC(1) = cViolet : catC(2) = cOrange
        For pt = 1 To 3
            .SeriesCollection(1).Points(pt).Interior.Color = catC(pt - 1)
            .SeriesCollection(1).Points(pt).Format.Fill.ForeColor.RGB = catC(pt - 1)
        Next pt
        .SeriesCollection(1).HasDataLabels = True
        With .SeriesCollection(1).DataLabels
            .ShowValue = True : .NumberFormat = "#,##0"
            .Position = xlLabelPositionOutsideEnd
            With .Font : .Color = cWhite : .Size = 8 : .Name = "Segoe UI" : End With
        End With
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .ChartGroups(1).GapWidth = 60
        On Error GoTo 0
    End With

    ' CHART 6 - Sales by Region (Column) - Bottom Right
    With wsDash.ChartObjects.Add(c3L, bT, w3, hB) : Set cht = .Chart : End With
    With cht
        .SetSourceData Source:=wsPivot.Range("R3:S7"), PlotBy:=xlColumns
        .ChartType = xlColumnClustered : .HasTitle = True
        ApplyTitleStyle .ChartTitle, "Sales by Region"
        ApplyChartAreaStyle cht : StyleAxes cht : AddValueGridlines cht
        HideFieldButtons cht : cht.HasLegend = False
        On Error Resume Next
        Dim regC(3) As Long
        regC(0) = cPurple : regC(1) = cTeal : regC(2) = cOrange : regC(3) = cCoral
        For pt = 1 To 4
            .SeriesCollection(1).Points(pt).Interior.Color = regC(pt - 1)
            .SeriesCollection(1).Points(pt).Format.Fill.ForeColor.RGB = regC(pt - 1)
        Next pt
        .SeriesCollection(1).HasDataLabels = True
        With .SeriesCollection(1).DataLabels
            .ShowValue = True : .NumberFormat = "#,##0"
            .Position = xlLabelPositionOutsideEnd
            With .Font : .Color = cWhite : .Size = 8 : .Bold = True : .Name = "Segoe UI" : End With
        End With
        .Axes(xlValue).MaximumScaleIsAuto = True
        .Axes(xlValue).MinimumScaleIsAuto = True
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .ChartGroups(1).GapWidth = 80
        On Error GoTo 0
    End With

End Sub
