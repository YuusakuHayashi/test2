Public Class DBTestViewModel
    'Inherits BaseViewModel2(Of DBTestModel)
    Inherits BaseViewModel2

    Public Overrides ReadOnly Property FrameType As String
        Get
            Return ViewModel.MAIN_FRAME
        End Get
    End Property


    Public Overrides Sub Initialize(ByRef m As Model, ByRef vm As ViewModel, ByRef adm As AppDirectoryModel, ByRef pim As ProjectInfoModel)
        'Throw New NotImplementedException()
    End Sub
End Class
