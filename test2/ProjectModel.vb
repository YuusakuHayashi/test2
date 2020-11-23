Public MustInherit Class ProjectModel
    Public MustOverride ReadOnly Property IconFileName As String
    Public MustOverride Function ViewSetupExecute(ByRef app As AppDirectoryModel,
                                                  ByRef vm As ViewModel) As ViewItemModel
    Public MustOverride Function ViewDefineExecute(ByVal modelname As String) As Object
End Class
