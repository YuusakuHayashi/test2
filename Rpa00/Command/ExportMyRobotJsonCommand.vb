Public Class ExportMyRobotJsonCommand : Inherits RpaCommandBase

    Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        If dat.Project.MyRobotObject Is Nothing Then
            Console.WriteLine($"'MyRobotObject' のインスタンスが設定されていません")
            Return False
        End If
        Return True
    End Function

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        dat.Project.MyRobotObject.Save(dat.Project.MyRobotJsonFileName, dat.Project.MyRobotObject)
        Console.WriteLine($"ファイル '{dat.Project.MyRobotJsonFileName}' にエクスポートしました")
        Console.WriteLine()
        Return 0
    End Function
End Class
