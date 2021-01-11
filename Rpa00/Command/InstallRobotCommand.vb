' 開発保留中
Public Class InstallRobotCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        Return MyBase.CanExecute(dat)
    End Function

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Throw New NotImplementedException()
    End Function
End Class
