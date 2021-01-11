Public Class RunRobotCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
        Get
            Return {0, 999}
        End Get
    End Property

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Dim i = 0
        dat.Project.MyRobotObject.Data = dat
        If dat.Project.CanExecute(dat) Then
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
End Class
