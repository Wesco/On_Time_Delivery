Attribute VB_Name = "Import_Sheets"
Option Explicit

Function ImportSheets() As Boolean
    Dim Path As String

    Path = Application.GetOpenFilename

    If Path <> "False" Then
        Workbooks.Open Path
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Drop In").Range("A1")
        ActiveWorkbook.Close
        ImportSheets = True
    Else
        ImportSheets = False
    End If

End Function
