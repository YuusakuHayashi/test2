Public Class RunRobotCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecutableProjectArchitectures As Integer()
        Get
            Return {
                RpaCodes.ProjectArchitecture.ClientServer,
                RpaCodes.ProjectArchitecture.IntranetClientServer,
                RpaCodes.ProjectArchitecture.StandAlone
            }
        End Get
    End Property

    Public Overrides ReadOnly Property CanExecute(trn As RpaTransaction, rpa As Object) As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableParameterCount() As Integer()
        Get
            Return {0, 99}
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableUserLevel() As Integer
        Get
            Return RpaCodes.ProjectUserLevel.User
        End Get
    End Property

    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object) As Integer
        Dim i = 0
        Call rpa.MyProjectObject.SetData(trn, rpa)
        If rpa.MyProjectObject.CanExecute() Then
            i = rpa.MyProjectObject.Execute()
        Else
            Console.WriteLine("ロボットの起動条件を満たしていません")
            i = 1000
        End If
        Return i
    End Function
End Class
