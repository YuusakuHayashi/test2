Public Class DelegateEventListener
    Private Shared _Instance As DelegateEventListener = New DelegateEventListener

    Public Shared ReadOnly Property Instance As DelegateEventListener
        Get
            Return DelegateEventListener._Instance
        End Get
    End Property

    Public Event ViewsChanged As EventHandler                   ' ビューエクスプローラー関連
    Public Event ViewResized As EventHandler
    Public Event TabViewClosed As EventHandler
    Public Event ProjectsChanged As EventHandler
    Public Event DeleteViewRequested As EventHandler
    Public Event OpenProjectRequested As EventHandler
    Public Event ViewSizeChanged As EventHandler
    Public Event RemoveFixedProjectRequested As EventHandler
    Public Event FixProjectRequested As EventHandler
    Public Event OpenViewRequested As EventHandler
    Public Event TabCloseRequested As EventHandler
    Public Event DeleteRequested As EventHandler
    Public Event ChildrenUpdated As EventHandler
    Public Event MigrateConditionUpdated As EventHandler
    Public Event MigrateConditionDeleteRequested As EventHandler
    Public Event DataTableCheckChanged As EventHandler

    Public Overloads Sub RaiseViewResized()
        RaiseEvent ViewResized(Nothing , EventArgs.Empty)
    End Sub

    Public Overloads Sub RaiseViewsChanged()
        RaiseEvent ViewsChanged(Nothing, EventArgs.Empty)
    End Sub

    Public Sub RaiseProjectsChanged()
        RaiseEvent ProjectsChanged(Nothing, EventArgs.Empty)
    End Sub

    Public Sub RaiseTabViewClosed()
        RaiseEvent TabViewClosed(Nothing, EventArgs.Empty)
    End Sub

    Public Sub RaiseDeleteViewRequested(ByVal [view] As ViewItemModel)
        RaiseEvent DeleteViewRequested([view], EventArgs.Empty)
    End Sub

    Public Sub RaiseOpenProjectRequested(ByVal project As ProjectInfoModel)
        RaiseEvent OpenProjectRequested(project, EventArgs.Empty)
    End Sub

    'Public Sub RaiseMultiViewSizeChanged(ByVal sender As Object)
    '    RaiseEvent MultiViewSizeChanged(sender, EventArgs.Empty)
    'End Sub
    'Public Sub RaiseOpenViewRequested(ByVal child As ViewItemModel)
    '    RaiseEvent OpenViewRequested(child, EventArgs.Empty)
    'End Sub

    Public Sub RaiseRemoveFixedProjectRequested(ByVal p As ProjectInfoModel)
        RaiseEvent RemoveFixedProjectRequested(p, EventArgs.Empty)
    End Sub

    Public Sub RaiseFixProjectRequested(ByVal p As ProjectInfoModel)
        RaiseEvent FixProjectRequested(p, EventArgs.Empty)
    End Sub

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
