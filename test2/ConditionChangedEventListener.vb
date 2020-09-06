Public Class ConditionChangedEventListener
    Private Shared _Instance As ConditionChangedEventListener = New ConditionChangedEventListener

    Public Shared ReadOnly Property Instance As ConditionChangedEventListener
        Get
            Return ConditionChangedEventListener._Instance
        End Get
    End Property

    Public Event DeleteRequested As EventHandler
    Public Event ChildrenUpdated As EventHandler
    Public Event MigrateConditionUpdated As EventHandler
    Public Event MigrateConditionDeleteRequested As EventHandler

    Public Sub RaiseDeleteRequested(ByVal child As ConditionModel)
        RaiseEvent DeleteRequested(child, EventArgs.Empty)
    End Sub

    Public Sub RaiseChildrenUpdated(ByVal child As ConditionModel)
        RaiseEvent ChildrenUpdated(child, EventArgs.Empty)
    End Sub

    Public Sub RaiseMigrateConditionUpdated(ByVal bro As MigrateModel)
        RaiseEvent MigrateConditionUpdated(bro, EventArgs.Empty)
    End Sub

    Public Sub RaiseMigrateConditionDeleteRequested(ByVal bro As MigrateModel)
        RaiseEvent MigrateConditionDeleteRequested(bro, EventArgs.Empty)
    End Sub

    Sub New()
    End Sub
End Class
