Public Class FileManagerModel
    Inherits ProjectModel(Of FileManagerModel)

    Private _LatestSettingFileName As String
    Public Property LatestSettingFileName As String
        Get
            Return Me._LatestSettingFileName
        End Get
        Set(value As String)
            Me._LatestSettingFileName = value
        End Set
    End Property

    '-------------------------------------------------------------------
    Public ReadOnly Property DefaultModelJson As String
        Get
            Return ProjectFile("defaultModel.json")
        End Get
    End Property

    Private _CurrentModelJson As String
    Public Property CurrentModelJson As String
        Get
            Return Me._CurrentModelJson
        End Get
        Set(value As String)
            Me._CurrentModelJson = value
        End Set
    End Property
    '-------------------------------------------------------------------

    '-------------------------------------------------------------------
    Public ReadOnly Property DefaultInitViewModelJson As String
        Get
            Return ProjectFile("defaultInitViewModel.json")
        End Get
    End Property

    Private _CurrentInitViewModelJson As String
    Public Property CurrentInitViewModelJson As String
        Get
            Return Me._CurrentInitViewModelJson
        End Get
        Set(value As String)
            Me._CurrentInitViewModelJson = value
        End Set
    End Property
    '-------------------------------------------------------------------

    '-------------------------------------------------------------------
    Public ReadOnly Property DefaultDBExplorerViewModelJson As String
        Get
            Return ProjectFile("defaultDBExplorerViewModel.json")
        End Get
    End Property

    Private _CurrentDBExplorerViewModelJson As String
    Public Property CurrentDBExplorerViewModelJson As String
        Get
            Return Me._CurrentDBExplorerViewModelJson
        End Get
        Set(value As String)
            Me._CurrentDBExplorerViewModelJson = value
        End Set
    End Property
    '-------------------------------------------------------------------
End Class
