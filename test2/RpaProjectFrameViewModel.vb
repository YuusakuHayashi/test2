Public Class RpaProjectFrameViewModel : Inherits BaseViewModel2

    Private _Menu As RpaProjectMenuViewModel
    Public Property Menu As RpaProjectMenuViewModel
        Get
            Return Me._Menu
        End Get
        Set(value As RpaProjectMenuViewModel)
            Me._Menu = value
        End Set
    End Property

    Private _Tabs As TabViewModel
    Public Property Tabs As TabViewModel
        Get
            Return Me._Tabs
        End Get
        Set(value As TabViewModel)
            Me._Tabs = value
        End Set
    End Property

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        Dim rpapmvm = New RpaProjectMenuViewModel
        Dim rpapvm = New RpaProjectViewModel
        Dim rpavm As RpaViewModel
        Dim tvm = New TabViewModel

        Call rpapmvm.Initialize(app, vm)
        Call rpapvm.Initialize(app, vm)

        tvm.AddTab(New ViewItemModel With {
            .Name = "RpaProject",
            .Content = rpapvm
        })

        Me.Menu = rpapmvm
        Me.Tabs = tvm
    End Sub
End Class
