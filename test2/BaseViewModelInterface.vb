Public Interface BaseViewModelInterface
    ReadOnly Property FrameType As String

    Sub Initialize(ByRef m As Model,
                   ByRef vm As ViewModel,
                   ByRef adm As AppDirectoryModel,
                   ByRef pim As ProjectInfoModel)
End Interface
