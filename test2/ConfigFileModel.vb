Public Class ConfigFileModel
    '-- Sql Status ---------------------------------------------------------------'
    Public SqlStatus As SqlModel
    '-----------------------------------------------------------------------------'


    '-- Tree View ----------------------------------------------------------------'
    Private _TreeViews As List(Of TreeViewModel)
    Public Property TreeViews As List(Of TreeViewModel)
        Get
            Return _TreeViews
        End Get
        Set(value As List(Of TreeViewModel))
            _TreeViews = value
        End Set
    End Property
    '-----------------------------------------------------------------------------'
End Class
