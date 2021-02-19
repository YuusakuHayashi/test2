Imports System.Windows.Forms
Imports System.IO

Module DspemuAutomation
    Sub Main()
        Dim exef As String = Application.ExecutablePath
        Dim rpaf As String = $"{Path.GetDirectoryName(exef)}\dspemuautomation.json"
        Dim errf As String = $"{Path.GetDirectoryName(exef)}\error"
        Dim chkf As String = $"{Path.GetDirectoryName(exef)}\check"
        Dim rpa As Rpa11.Rpa11 = Rpa11.Rpa11.LoadFromFile(rpaf)

        If File.Exists(errf) Then
            File.Delete(errf)
        End If
        If File.Exists(chkf) Then
            File.Delete(chkf)
        End If

        If rpa IsNot Nothing Then
            If rpa.CheckRun() Then
                rpa.Run()
            Else
                File.Create(chkf)
            End If
        Else
            File.Create(errf)
        End If
    End Sub
End Module
