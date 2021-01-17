Imports Rpa00

Public Class AllSaveCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Call dat.Transaction.SaveCommandLogs()
        Call RpaModule.Save(RpaCui.SystemIniFileName, dat.Initializer, RpaInitializer.SystemIniChangedFileName)
        Call RpaModule.Save(dat.Project.SystemJsonFileName, dat.Project, dat.Project.SystemJsonChangedFileName)
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
