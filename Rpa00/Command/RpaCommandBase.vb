Public MustInherit Class RpaCommandBase
    Public MustOverride Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer

    Public Overridable ReadOnly Property ExecutableProjectArchitectures As Integer()
        Get
            Return {
                RpaCodes.ProjectArchitecture.ClientServer,
                RpaCodes.ProjectArchitecture.IntranetClientServer,
                RpaCodes.ProjectArchitecture.StandAlone
            }
        End Get
    End Property

    Public Overridable ReadOnly Property CanExecute(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overridable ReadOnly Property ExecutableParameterCount() As Integer()
        Get
            Return {0, 0}
        End Get
    End Property

    Public Overridable ReadOnly Property ExecutableUserLevel() As Integer
        Get
            Return RpaCodes.ProjectUserLevel.User
        End Get
    End Property
End Class
