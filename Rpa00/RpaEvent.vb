Public Class RpaEvent
    Private Shared _Instance As RpaEvent = New RpaEvent
    Public Shared ReadOnly Property Instance As RpaEvent
        Get
            Return RpaEvent._Instance
        End Get
    End Property

    Public Event ProjectChanged As EventHandler
    Public Sub RaiseProjectChanged(ByVal pname As String)
        RaiseEvent ProjectChanged(Me, New ProjectChangedEventArgs With {.PropertyName = pname})
    End Sub

    Public Class ProjectChangedEventArgs : Inherits EventArgs
        Private _PropertyName As String
        Public Property PropertyName As String
            Get
                Return Me._PropertyName
            End Get
            Set(value As String)
                Me._PropertyName = value
            End Set
        End Property
    End Class
End Class
