Public Class DelegateEventListener
    Private Shared _Instance As DelegateEventListener = New DelegateEventListener

    Public Shared ReadOnly Property Instance As DelegateEventListener
        Get
            Return DelegateEventListener._Instance
        End Get
    End Property

    Public Event OpenViewRequested As EventHandler
    Public Event TabCloseRequested As EventHandler
    Public Event DeleteRequested As EventHandler
    Public Event ChildrenUpdated As EventHandler
    Public Event MigrateConditionUpdated As EventHandler
    Public Event MigrateConditionDeleteRequested As EventHandler
    Public Event DataTableCheckChanged As EventHandler

    'Public Sub RaiseOpenViewRequested(ByVal child As ViewItemModel)
    '    RaiseEvent OpenViewRequested(child, EventArgs.Empty)
    'End Sub
    Public Sub RaiseOpenViewRequested(ByVal v As ViewItemModel)
        RaiseEvent OpenViewRequested(v, EventArgs.Empty)
    End Sub

    Public Sub RaiseTabCloseRequested(ByVal child As TabItemModel)
        RaiseEvent TabCloseRequested(child, EventArgs.Empty)
    End Sub

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
