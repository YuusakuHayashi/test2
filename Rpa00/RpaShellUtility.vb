Public Class RpaShellUtility : Inherits RpaUtilityBase
    Public Sub RunShell(ByVal exe As String, ByVal arg As String)
        Dim proc = New System.Diagnostics.Process()
        proc.StartInfo.FileName = exe
        proc.StartInfo.UseShellExecute = False
        proc.StartInfo.RedirectStandardOutput = True
        proc.StartInfo.RedirectStandardInput = False
        proc.StartInfo.CreateNoWindow = True
        proc.StartInfo.Arguments = arg
        proc.Start()
        proc.WaitForExit()
        proc.Close()
    End Sub
End Class
