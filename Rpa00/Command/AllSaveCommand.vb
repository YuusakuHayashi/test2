Imports Rpa00

Public Class AllSaveCommand : Inherits RpaCommandBase
    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Call RpaModule.Save(RpaCui.SystemIniFileName, dat.Initializer, RpaInitializer.SystemIniChangedFileName)
        Call RpaModule.Save(dat.Project.SystemJsonFileName, dat.Project, dat.Project.SystemJsonChangedFileName)
        Return 0
    End Function
End Class
