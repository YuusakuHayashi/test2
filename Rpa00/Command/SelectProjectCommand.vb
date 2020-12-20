Public Class SelectProjectCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecutableProjectArchitectures As Integer()
        Get
            Return {
                RpaCodes.ProjectArchitecture.ClientServer,
                RpaCodes.ProjectArchitecture.StandAlone,
                RpaCodes.ProjectArchitecture.IntranetClientServer
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
            Return {1, 1}
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableUserLevel() As Integer
        Get
            Return RpaCodes.ProjectUserLevel.User
        End Get
    End Property

    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object) As Integer
        If trn.Parameters.Count = 0 Then
            Console.WriteLine($"プロジェクトが選択されていません")
            Return 1000
        Else
            'rpa = New T With {.ProjectName = trn.Parameters(0)}
            Return 0
        End If
    End Function
End Class
