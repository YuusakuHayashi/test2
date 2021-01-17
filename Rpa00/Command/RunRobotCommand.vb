Public Class RunRobotCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim i = 0
        dat.Project.MyRobotObject.Data = dat
        Console.WriteLine()
        Console.WriteLine($"プロジェクトの起動条件を確認しています・・・")
        If dat.Project.CanExecute(dat) Then
            Console.WriteLine($"ロボットの起動条件を確認しています・・・")
            If dat.Project.MyRobotObject.CanExecute(dat) Then
                i = dat.Project.MyRobotObject.Execute(dat)
            Else
                Console.WriteLine("ロボットが起動条件を満たしていません")
                Console.WriteLine()
                i = 1000
            End If
        Else
            Console.WriteLine("プロジェクトが起動条件を満たしていません")
            Console.WriteLine()
            i = 1000
        End If
        Return i
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Me.Main
        Me.ExecutableParameterCount = {0, 999}
    End Sub
End Class
