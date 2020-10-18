Public Class DBTestViewModel
    'Inherits BaseViewModel2(Of DBTestModel)
    Inherits BaseViewModel2

    Public Overrides ReadOnly Property ViewType As String
        Get
            Return ViewModel.MAIN_VIEW
        End Get
    End Property
End Class
