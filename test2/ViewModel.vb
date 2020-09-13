Public Class ViewModel
    Inherits BaseViewModel

    Private _ContextModel As Object
    Public Property ContextModel As Object
        Get
            Return Me._ContextModel
        End Get
        Set(value As Object)
            Me._ContextModel = value
            RaisePropertyChanged("ContextModel")
        End Set
    End Property

    'Private _ContextDistination As String
    'Public Property ContextDistination As String
    '    Get
    '        Return Me._ContextDistination
    '    End Get
    '    Set(value As String)
    '        Me._ContextDistination = value
    '    End Set
    'End Property
End Class
