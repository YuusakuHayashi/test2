Public Class ViewModelChangedEventListener
    Private Shared _Instance As ViewModelChangedEventListener = New ViewModelChangedEventListener

    Public Shared ReadOnly Property Instance As ViewModelChangedEventListener
        Get
            Return ViewModelChangedEventListener._Instance
        End Get
    End Property

    Public Event ContextChangeRequested As EventHandler

    Public Sub RaiseContextChangeRequested(ByRef m As Model, ByRef vm As ViewModel)
        RaiseEvent ContextChangeRequested(m, EventArgs.Empty)
    End Sub
End Class
