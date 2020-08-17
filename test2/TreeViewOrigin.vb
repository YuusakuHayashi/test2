Public Class TreeViewOrigin
    Private Property _RealName As String
    Public Property RealName As String
        Get
            Return _RealName
        End Get
        Set(value As String)
            _RealName = value
        End Set
    End Property

    Private _Children As List(Of TreeViewOrigin)
    Public Property Children As List(Of TreeViewOrigin)
        Get
            Return _Children
        End Get
        Set(value As List(Of TreeViewOrigin))
            _Children = value
        End Set
    End Property
End Class
