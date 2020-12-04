For i = 0 To (WScript.Arguments.Count - 1)
    If WScript.Arguments(i) = "Updater.bas" Then
        If WScript.Arguments.Count > 1 Then
            WScript.StdOut.WriteLine "update canceled because your args include 'Updater.bas'"
            WScript.Quit
        End If
    End If
Next

'--- このように指定すると、Str型として使おうとしたとき、スクリプトの実行ディレクトリをフルパスで返す---'
'Set pth = fso.GetFolder(".")
'------------------------------------------------------------------------------------------------------'


If WScript.Arguments(1) = "mac_importer.bas" Then
    Call update_mac("wrk.xlsm", "mac_importer.main2")
    Call update_mac("rpa.xlsm", "mac_importer.main2")
Else
    Call update_mac("rpa.xlsm", "mac_importer.main")
End If

' SUB ROUTINE PROCEDURE
'--------------------------------------------------------------------------------------------------'
'Private Sub update_1 ()
'    On Error Resume Next
'        Set exapp = exobj.Workbooks.Open(pth + "\Updater.xlsm")
'        exobj.Application.Run _
'            "Updater.main", _
'            SOURCE_DIR, _
'            pth + "\Main.xlsm", _
'            WScript.Arguments(0)
'    On Error GoTo 0
'    If Err.Number <> 0 Then
'        WScript.StdOut.WriteLine "Update Error"
'    End If
'    exapp.Save
'    exapp.Close
'
'    On Error Resume Next
'        Set exapp = exobj.Workbooks.Open(pth + "\Main.xlsm")
'        exobj.Application.Run _
'            "Updater.main", _
'            SOURCE_DIR, _
'            pth + "\Updater.xlsm", _
'            WScript.Arguments(0)
'    On Error GoTo 0
'    If Err.Number <> 0 Then
'        WScript.StdOut.WriteLine "Update Error"
'    End If
'    exapp.Save
'    exapp.Close
'End Sub
'--------------------------------------------------------------------------------------------------'

'--------------------------------------------------------------------------------------------------'
Private Sub update_mac (xlm, methd)
    Dim fso
    Dim exapp
    Dim wbook
    Dim imp_mod
    Dim rem_mod
    Dim pth
    Dim i

    Set fso   = CreateObject("Scripting.FileSystemObject")
    Set exapp = CreateObject("Excel.Application")
    exapp.Visible = False
    pth = WScript.Arguments(0)

    Set wbook = exapp.Workbooks.Open(pth + "\" + xlm, 0)
    wbook.Application.DisplayAlerts = False
    wbook.Application.EnableEvents = False
    On Error Resume Next
        If WScript.Arguments.Count > 0 Then
            For i = 1 To (WScript.Arguments.Count - 1)
                If fso.FileExists(pth + "\" + WScript.Arguments(i)) Then
                    Select Case fso.GetExtensionName(WScript.Arguments(i))
                        Case "bas", "cls"
                            rem_mod = fso.GetBaseName(WScript.Arguments(i))
                            imp_mod = WScript.Arguments(i)
                       	    exapp.Application.Run methd, imp_mod, rem_mod
                    End Select
                Else
                    WScript.StdOut.WriteLine "File " + pth + " " + WScript.Arguments(i) + "  isnt in Server"
                End If
            Next
        End If
    On Error GoTo 0
    If Err.Number <> 0 Then
        WScript.StdOut.WriteLine "Update Error"
    End If
    wbook.Save
    wbook.Close
    exapp.Quit

    Set fso   = Nothing
    Set exapp = Nothing
    Set wbook = Nothing
End Sub
'--------------------------------------------------------------------------------------------------'

