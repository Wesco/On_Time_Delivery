Attribute VB_Name = "Program"
Option Explicit

Sub FH_Main()
    Application.ScreenUpdating = False
    
    If ImportSheets = True Then
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
