Public Class RpaInitializer : Inherits JsonHandler(Of RpaInitializer)
    Private _CurrentProjectArchitecture As String
    Public Property CurrentProjectArchitecture As String
        Get
            Return Me._CurrentProjectArchitecture
        End Get
        Set(value As String)
            Me._CurrentProjectArchitecture = value
        End Set
    End Property

    Private _AutoLoad As Boolean
    Public Property AutoLoad As Boolean
        Get
            Return Me._AutoLoad
        End Get
        Set(value As Boolean)
            Me._AutoLoad = value
        End Set
    End Property
End Class
