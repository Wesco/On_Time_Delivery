Attribute VB_Name = "Program"
Option Explicit
Public Const RepositoryName As String = "On_Time_Delivery"
Public Const VersionNumber = "1.0.1"

Public FHType As String

Sub FH_Main()
    Application.ScreenUpdating = False

    If ImportSheets = True Then
        FHType = Sheets("Drop In").Range("J2").Value
        RemoveData
        ExportDropIn
    End If

    Application.ScreenUpdating = True
End Sub

Sub NU_Main()
    Dim w As Variant

    For Each w In Workbooks
        If InStr(w.Name, "Integrated Supply POLineReport") Then
            w.Activate
        End If
    Next
    FormatSheet
End Sub

Sub Clean()
    Dim s As Worksheet

    Application.DisplayAlerts = False

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            Cells.Delete
            Range("A1").Select
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select

    Application.DisplayAlerts = True
End Sub
