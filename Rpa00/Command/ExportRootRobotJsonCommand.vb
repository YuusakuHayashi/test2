Imports Rpa00

Public Class ExportRootRobotJsonCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecutableProjectArchitectures As String()
        Get
            Return {(New IntranetClientServerProject).GetType.Name}
        End Get
    End Property

    Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        If dat.Project.RootRobotObject Is Nothing Then
            Console.WriteLine($"'RootRobotObject' のインスタンスが設定されていません")
            Return False
        End If
        Return True
    End Function

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        dat.Project.RootRobotObject.Save(dat.Project.RootRobotJsonFileName, dat.Project.RootRobotObject)
        Console.WriteLine($"ファイル '{dat.Project.RootRobotJsonFileName}' にエクスポートしました")
        Console.WriteLine()
        Return 0
    End Function
End Class
