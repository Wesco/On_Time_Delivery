Attribute VB_Name = "Format_Data"
Option Explicit

Sub RemoveData()
    Sheets("Drop In").Select
    Rows(1).Delete
    FilterSheet "WESCO DISTRIBUTION INC", 23
    FilterSheet "Late", 25
    FilterDate
    
    Columns("V").Delete
    Columns("S").Delete
    Columns("J:P").Delete
    Columns("H").Delete
    Columns("D").Delete
End Sub

Sub FilterDate()
    Dim Rng As Range
    Dim aRng() As Variant
    Dim aHeaders As Variant
    Dim iCounter As Long
    Dim i As Long
    Dim y As Long

    Set Rng = ActiveSheet.UsedRange
    aHeaders = Range(Cells(1, 1), Cells(1, ActiveSheet.UsedRange.Columns.Count))
    iCounter = 1

    Do While iCounter <= Rng.Rows.Count
        If Format(Rng(iCounter, 18).Value, "yymm") = Format(Date, "yymm") Then
            i = i + 1
        End If
        iCounter = iCounter + 1
    Loop

    ReDim aRng(1 To i, 1 To Rng.Columns.Count) As Variant

    iCounter = 1
    i = 0
    Do While iCounter <= Rng.Rows.Count
        If Format(Rng(iCounter, 18).Value, "yymm") = Format(Date, "yymm") Then
            i = i + 1
            For y = 1 To Rng.Columns.Count
                aRng(i, y) = Rng(iCounter, y)
            Next
        End If
        iCounter = iCounter + 1
    Loop

    ActiveSheet.Cells.Delete
    Range(Cells(1, 1), Cells(UBound(aRng, 1), UBound(aRng, 2))) = aRng
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, UBound(aHeaders, 2))) = aHeaders

End Sub

Sub FormatSheet()
    Dim iRows As Long
    Dim iCols As Integer

    ActiveSheet.UsedRange.UnMerge
    Rows("1:5").Delete
    Columns("V:V").Delete
    Columns("R:R").Delete
    Columns("Q:Q").Delete
    Columns("G:G").Delete
    Columns("F:F").Delete
    Columns("C:C").Delete
    Columns("A:A").Delete

    iCols = ActiveSheet.UsedRange.Columns.Count
    iRows = ActiveSheet.UsedRange.Rows.Count

    ActiveWorkbook.Worksheets("Sheet1").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").sort.SortFields.Add Key:=Range("J1"), _
                                                            SortOn:=xlSortOnValues, _
                                                            Order:=xlAscending, _
                                                            DataOption:=xlSortNormal

    With ActiveWorkbook.Worksheets("Sheet1").sort
        .SetRange Range(Cells(2, 1), Cells(iRows, iCols))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


    Range("K1").Select
    ActiveWorkbook.Worksheets("Sheet1").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").sort.SortFields.Add(Range(Cells(2, 11), Cells(iRows, 11)), _
                                                            xlSortOnCellColor, _
                                                            xlAscending, , _
                                                            xlSortNormal _
                                                            ).SortOnValue.Color = RGB(255, 0, 0)

    ActiveWorkbook.Worksheets("Sheet1").sort.SortFields.Add(Range(Cells(2, 11), Cells(iRows, 11)), _
                                                            xlSortOnCellColor, _
                                                            xlAscending, , _
                                                            xlSortNormal _
                                                            ).SortOnValue.Color = RGB(255, 255, 0)

    ActiveWorkbook.Worksheets("Sheet1").sort.SortFields.Add(Range(Cells(2, 11), Cells(iRows, 11)), _
                                                            xlSortOnCellColor, _
                                                            xlAscending, , _
                                                            xlSortNormal _
                                                            ).SortOnValue.Color = RGB(0, 255, 0)

    ActiveWorkbook.Worksheets("Sheet1").sort.SortFields.Add(Range(Cells(2, 11), Cells(iRows, 11)), _
                                                            xlSortOnCellColor, _
                                                            xlAscending, , _
                                                            xlSortNormal _
                                                            ).SortOnValue.Color = RGB(255, 255, 255)

    ActiveWorkbook.Worksheets("Sheet1").sort.SortFields.Add(Range(Cells(2, 11), Cells(iRows, 11)), _
                                                            xlSortOnCellColor, _
                                                            xlAscending, , _
                                                            xlSortNormal _
                                                            ).SortOnValue.Color = RGB(0, 0, 255)
    With ActiveWorkbook.Worksheets("Sheet1").sort
        .SetRange Range(Cells(1, 1), Cells(iRows, iCols))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").AutoFilter
End Sub
