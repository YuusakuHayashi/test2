Public Class UpdateMyRobotCommand : Inherits RpaCommandBase

    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        If dat.Project.MyRobotObject Is Nothing Then
            Console.WriteLine($"'MyRobotObject' のインスタンスが設定されていません")
            Return False
        End If
        Return True
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        dat.Project.MyRobotObject.Save(dat.Project.MyRobotJsonFileName, dat.Project.MyRobotObject)
        Console.WriteLine($"ファイル '{dat.Project.MyRobotJsonFileName}' にエクスポートしました")
        Console.WriteLine()
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
    End Sub
End Class
