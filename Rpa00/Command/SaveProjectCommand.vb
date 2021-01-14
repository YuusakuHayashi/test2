Public Class SaveProjectCommand : Inherits RpaCommandBase

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Call RpaModule.Save(dat.Project.SystemJsonFileName, dat.Project, dat.Project.SystemJsonChangedFileName)
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
