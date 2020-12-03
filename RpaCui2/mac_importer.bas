Attribute VB_Name = "mac_importer"
Option Explicit


'README


Private Sub main(ParamArray pa() As Variant)
    Dim imp_mod As String: imp_mod = pa(0)
    Dim rem_mod As String: rem_mod = pa(1)

    With ThisWorkbook.VBProject
        On Error Resume Next
            .VBComponents.Remove .VBComponents(rem_mod)
        On Error GoTo 0
        If Err.Number > 0 And Err.Number <> 9 Then
            Err.Raise (Err.Number)
            Exit Sub
        End If
        .VBComponents.Import ThisWorkbook.Path + "\" + imp_mod 
    End With
    'Set fso = Nothing
End Sub

Private Sub main2(ParamArray pa() As Variant)
    Dim imp_mod As String: imp_mod = pa(0)
    Dim rem_mod As String: rem_mod = pa(1)
    Dim wbook As Workbook

    Dim dst As String
    Select Case ThisWorkbook.Name
        Case "rpa.xlsm"
            dst = "wrk.xlsm"
        Case "wrk.xlsm"
            dst = "rpa.xlsm"
    End Select

    Set wbook = Workbooks.Open(ThisWorkbook.Path + "\" + dst, 0)
    wbook.Application.DisplayAlerts = False
    wbook.Application.EnableEvents = False

    With wbook.VBProject
        On Error Resume Next
            .VBComponents.Remove .VBComponents(rem_mod)
        On Error GoTo 0
        If Err.Number > 0 And Err.Number <> 9 Then
            Err.Raise (Err.Number)
            Exit Sub
        End If
        .VBComponents.Import ThisWorkbook.Path + "\" + imp_mod 
    End With

    wbook.Save
    wbook.Close
    Set wbook = Nothing
End Sub
