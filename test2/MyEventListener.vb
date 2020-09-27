Public Class MyEventListener
    Private Shared _Instance As MyEventListener = New MyEventListener

    Public Shared ReadOnly Property Instance As MyEventListener
        Get
            Return MyEventListener._Instance
        End Get
    End Property

    Public Event DeleteRequested As EventHandler
    Public Event ChildrenUpdated As EventHandler
    Public Event MigrateConditionUpdated As EventHandler
    Public Event MigrateConditionDeleteRequested As EventHandler
    Public Event DataTableCheckChanged As EventHandler

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

    Public Sub RaiseDataTableCheckChanged()
        RaiseEvent DataTableCheckChanged(Nothing, EventArgs.Empty)
    End Sub
End Class
