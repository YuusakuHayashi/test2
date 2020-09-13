Public Class FileContentsModel
    Protected Function ProjectFile(ByVal f As String) As String
        If f.Contains(Me.ProjectDirectoryName) Then
            ProjectFile = f
        Else
            ProjectFile = Me.ProjectDirectoryName & "\" & f
        End If
    End Function

    Protected Function _ProjectDirectory(ByVal d As String) As String
        _ProjectDirectory = ProjectFile(d)
    End Function

    Protected ReadOnly Property ProjectDirectoryName As String
        Get
            Return System.Environment.GetEnvironmentVariable("USERPROFILE") & "\test"
        End Get
    End Property

    Protected ReadOnly Property _ProjectIniFile As String
        Get
            Return Me.ProjectDirectoryName & "\ProjectInit"
        End Get
    End Property

    Public ReadOnly Property FileManagerJson As String
        Get
            Return Me.ProjectDirectoryName & "\FileManager.json"
        End Get
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

    Public ReadOnly Property DefaultTestModelJson As String
        Get
            Return ProjectFile("defaultTestModel.json")
        End Get
    End Property
    '-------------------------------------------------------------------

    'Public Overrides Sub MemberCheck()
    '    ' メンバのデフォルト値を設定したい場合、ここに記述
    '    If String.IsNullOrEmpty(Me.CurrentModelJson) Then
    '        Me.CurrentModelJson = Me.DefaultModelJson
    '    End If
    '    If String.IsNullOrEmpty(Me.CurrentInitViewModelJson) Then
    '        Me.CurrentInitViewModelJson = Me.DefaultInitViewModelJson
    '    End If
    '    If String.IsNullOrEmpty(Me.CurrentDBExplorerViewModelJson) Then
    '        Me.CurrentDBExplorerViewModelJson = Me.DefaultDBExplorerViewModelJson
    '    End If
    'End Sub

End Class
