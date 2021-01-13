Attribute VB_Name = "MacroImporter"
Option Explicit


'README

Private Sub Main(ByRef target() As Variant)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim mfile As String: mfile = target(0)
    Dim module As String: module = fso.GetBaseName(mfile)
    Dim extension As String: extension = fso.GetExtensionName(mfile)
    With ThisWorkbook.VBProject
        If extension = "bas" Or extension = "cls" Then
            On Error Resume Next
                Call .VBComponents.Remove(.VBComponents(module))
            On Error GoTo 0
            If Err.Number > 0 And Err.Number <> 9 Then
                Err.Raise (Err.Number)
                Exit Sub
            Else
                Call .VBComponents.Import(mfile)
            End If
        End If
    End With
    ThisWorkbook.Save
End Sub

Private Function IsModuleExist(ByRef v() As Variant) As Boolean
    Dim module As String: module = v(0)
    Dim ck As Boolean: ck = False
    Dim vbc As Object

    For Each vbc In ThisWorkbook.VBProject.VBComponents
        If vbc.Name = module Then
            ck = True
            Exit For
        End If
    Next

    If ck Then
        IsModuleExist = True
    Else
        IsModuleExist = False
    End If

    Set vbc = Nothing
End Function

Private Function ShowModules(ByRef v() As Variant) As String
    Dim vbc As Object
    Dim txt As String: txt = vbNullString

    For Each vbc In ThisWorkbook.VBProject.VBComponents
        txt = txt & vbc.Name & vbCrLf
    Next

    ShowModules = txt
    Set vbc = Nothing
End Function

'Private Sub main(ParamArray pa() As Variant)
'    Dim imp_mod As String: imp_mod = pa(0)
'    Dim rem_mod As String: rem_mod = pa(1)
'
'    With ThisWorkbook.VBProject
'        On Error Resume Next
'            .VBComponents.Remove .VBComponents(rem_mod)
'        On Error GoTo 0
'        If Err.Number > 0 And Err.Number <> 9 Then
'            Err.Raise (Err.Number)
'            Exit Sub
'        End If
'        .VBComponents.Import ThisWorkbook.Path + "\" + imp_mod
'    End With
'    'Set fso = Nothing
'End Sub

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
