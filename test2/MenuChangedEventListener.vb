Public Class MenuChangedEventListener
    Private Shared _Instance As MenuChangedEventListener = New MenuChangedEventListener

    Public Shared ReadOnly Property Instance As MenuChangedEventListener
        Get
            Return MenuChangedEventListener._Instance
        End Get
    End Property

    Public Event ChangeRequested As EventHandler

    Public Sub RaiseChangeRequested(ByVal child As _MenuModel)
        RaiseEvent ChangeRequested(child, EventArgs.Empty)
    End Sub
End Class
