Attribute VB_Name = "Export_Sheets"
Option Explicit

Sub ExportDropIn()
    Dim FileName As String
    Dim FilePath As String
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim TotalRows As Long

    Application.DisplayAlerts = False

    Sheets("Drop In").Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value

    'Filter sites
    ActiveSheet.UsedRange.AutoFilter Field:=1, _
                                     Criteria1:=Array("NF-BECKJ", "NF-MFORT", "NF-ZIMMR", "RC-EDWCC", "RC-MDWSC", "RC-NOBLE", _
                                                      "RF-CAYUG", "RF-EBEND", "RF-GALLA", "RF-GBSON", "RC-WABRP", "RF-WBASH", _
                                                      "NC-HROCK"), _
                                                      Operator:=xlFilterValues

    ActiveSheet.UsedRange.Copy Destination:=Sheets("Temp").Range("A1")

    'Remove filtered data
    Cells.Delete

    'Reinsert column headers
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders

    'Sort the filtered data
    SortFHData Sheets("Temp")

    'Fix date formats
    Sheets("Temp").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    Range("H2:K" & TotalRows).NumberFormat = "m/d/yyyy"
    Range("P2:P" & TotalRows).NumberFormat = "m/d/yyyy"

    'Copy the sheet
    Sheets("Temp").Copy
    ActiveSheet.Name = "Sheet1"

    'Prompt user to save
    If FHType = "STD" Then
        MsgBox "Please save the SPOT report."
    ElseIf FHType = "REL" Then
        MsgBox "Please save the REL report."
    End If

    If Application.Dialogs(xlDialogSaveAs).Show = True Then
        FilePath = ActiveWorkbook.Path
        FileName = ActiveWorkbook.Name
        ActiveWorkbook.Close

        'Email saved report
        If FHType = "STD" Then
            Email "rmason@wesco.com; vstrickland@wesco.com", Subject:="On-Time Delivery Report", Attachment:=FilePath & "\" & FileName
        ElseIf FHType = "REL" Then
            Email "kwitzman@wesco.com", Subject:="On-Time Delivery Report", Attachment:=FilePath & "\" & FileName
        Else
            MsgBox "Report type """ & FHType & """ not recognized. Email not sent."
        End If
    Else
        ActiveWorkbook.Close
    End If

    'Sort data
    SortFHData Sheets("Drop In")

    'Fix date formats
    Sheets("Drop In").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    Range("H2:K" & TotalRows).NumberFormat = "m/d/yyyy"
    Range("P2:P" & TotalRows).NumberFormat = "m/d/yyyy"

    MsgBox "Please save your report."
    ActiveSheet.Copy
    ActiveSheet.UsedRange.Columns.AutoFit
    ActiveSheet.Name = "Sheet1"

    'Prompt to user save
    Application.Dialogs(xlDialogSaveAs).Show

    'Close the macro
    ThisWorkbook.Close

    Application.DisplayAlerts = True
End Sub

Sub SortFHData(ws As Worksheet)
    Dim PrevSheet As Worksheet
    Set PrevSheet = ActiveSheet
    Dim TotalRows As Long

    ws.Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    ActiveSheet.sort.SortFields.Clear
    With ActiveSheet.sort
        'Add sort for PO Status CD
        .SortFields.Add Key:=Range("G2:G" & TotalRows), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal

        'Add sort for Req Delv Dt Adj
        .SortFields.Add Key:=Range("I2:I2" & TotalRows), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal

        'Apply the sort
        .SetRange ActiveSheet.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    PrevSheet.Select
End Sub
