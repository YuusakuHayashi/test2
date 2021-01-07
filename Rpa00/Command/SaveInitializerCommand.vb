Public Class SaveInitializerCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Call RpaModule.Save(RpaCui.SystemIniFileName, dat.Initializer, RpaInitializer.SystemIniChangedFileName)
        Return 0
    End Function
End Class
