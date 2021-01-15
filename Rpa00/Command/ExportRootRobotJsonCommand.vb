Imports Rpa00

Public Class ExportRootRobotJsonCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecutableProjectArchitectures As String()
        Get
            Return {(New IntranetClientServerProject).GetType.Name}
        End Get
    End Property

    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        If dat.Project.RootRobotObject Is Nothing Then
            Console.WriteLine($"'RootRobotObject' のインスタンスが設定されていません")
            Return False
        End If
        Return True
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        dat.Project.RootRobotObject.Save(dat.Project.RootRobotJsonFileName, dat.Project.RootRobotObject)
        Console.WriteLine($"ファイル '{dat.Project.RootRobotJsonFileName}' にエクスポートしました")
        Console.WriteLine()
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecutableUser = {"RootUser"}
    End Sub
End Class
