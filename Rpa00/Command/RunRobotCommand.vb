Public Class RunRobotCommand : Inherits RpaCommandBase
    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Dim i = 0
        Call dat.Project.MyProjectObject.SetData(dat)
        If dat.Project.MyProjectObject.CanExecute Then
            i = dat.Project.MyProjectObject.Execute()
        Else
            Console.WriteLine("ロボットの起動条件を満たしていません")
            i = 1000
        End If
        Return i
    End Function
End Class
