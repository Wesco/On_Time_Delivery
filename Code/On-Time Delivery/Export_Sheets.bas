Attribute VB_Name = "Export_Sheets"
Option Explicit

Sub ExportDropIn()
    Dim s As Worksheet
    
    Application.DisplayAlerts = False
    Dim sPath As String
    Sheets("Drop In").Copy
    Set s = ActiveSheet
    ActiveSheet.UsedRange.Columns.EntireColumn.AutoFit
    s.Name = "Sheet1"
    Application.Dialogs(xlDialogSaveAs).Show
    ThisWorkbook.Close
    Application.DisplayAlerts = True
End Sub
