Public Class SaveInitializerCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Call dat.System.Save(RpaCui.SystemIniFileName, dat.Initializer, RpaInitializer.SystemIniChangedFileName)
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecuteIfNoProject = True
    End Sub
End Class
