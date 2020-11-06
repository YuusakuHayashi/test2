Public MustInherit Class ProjectModel
    Public MustOverride Sub ViewSetupExecute(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
    Public MustOverride Function ViewDefineExecute(ByVal modelname As String) As Object
End Class
