Imports System.Windows.Forms
Imports System.IO

Module DspemuAutomation
    Sub Main()
        Dim exef As String = Application.ExecutablePath
        Dim rpaf As String = $"{Path.GetDirectoryName(exef)}\dspemuautomation.json"
        Try
            Dim rpa As Rpa11.Rpa11 = Rpa11.Rpa11.LoadFromFile(rpaf)
            If rpa IsNot Nothing Then
                If rpa.CheckRun() Then
                    rpa.PrintExecutingJob()
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
End Module
