Imports Rpa00

Public Class SelectProjectCommand : Inherits RpaCommandBase
    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        If trn.Parameters.Count = 0 Then
            Console.WriteLine($"プロジェクトが選択されていません")
            Return 1000
        Else
            rpa.ProjectName = trn.Parameters(0)
            Return 0
        End If
    End Function
End Class
