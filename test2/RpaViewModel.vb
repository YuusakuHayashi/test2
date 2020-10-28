Imports test2

Public Class RpaViewModel : Inherits BaseViewModel2
    Private _ProjectRootDirectoryName As String
    Public Property ProjectRootDirectoryName As String
        Get
            Return Me._ProjectRootDirectoryName
        End Get
        Set(value As String)
            Me._ProjectRootDirectoryName = value
        End Set
    End Property

    Private _ProjectUserDirectoryName As String
    Public Property ProjectUserDirectoryName As String
        Get
            Return Me._ProjectUserDirectoryName
        End Get
        Set(value As String)
            Me._ProjectUserDirectoryName = value
        End Set
    End Property

    Public Sub _ViewInitializing()
        With AppInfo.ProjectInfo.Model.Data
        End With
    End Sub

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        InitializeHandler = AddressOf _ViewInitializing
        Call BaseInitialize(app, vm)
    End Sub
End Class
