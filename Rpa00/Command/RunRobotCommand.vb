Public Class RunRobotCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim i = 0
        Dim rck As Boolean = True

        'MyRobotObjectのリロード用
        Dim j As Integer = dat.Project.SwitchRobot(dat.Project.RobotName)

        dat.Project.MyRobotObject.Data = dat
        Call dat.System.RpaWriteLine()
        Call dat.System.RpaWriteLine($"プロジェクトの起動条件を確認しています・・・")
        If dat.Project.CanExecute(dat) Then
            Call dat.System.RpaWriteLine($"ロボットの起動条件を確認しています・・・")
            If dat.Project.MyRobotObject.CanExecute(dat) Then
                i = dat.Project.MyRobotObject.Execute(dat)
            Else
                Call dat.System.RpaWriteLine("ロボットが起動条件を満たしていません")
                Call dat.System.RpaWriteLine()
                i = 1000
            End If
        Else
            Call dat.System.RpaWriteLine("プロジェクトが起動条件を満たしていません")
            Call dat.System.RpaWriteLine()
            i = 1000
        End If
        Return i
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Me.Main
        Me.ExecutableParameterCount = {0, 999}
    End Sub
End Class
