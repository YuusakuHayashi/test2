Imports System.Collections.ObjectModel

Public Class AppDirectoryModel
    Public ReadOnly Property AppDirectoryName As String
        Get
            Return System.Environment.GetEnvironmentVariable("USERPROFILE") & "\hys"
        End Get
    End Property

    'Private _CurrentProjects As ObservableCollection(Of ProjectModel)
    'Public Property CurrentProjects As ObservableCollection(Of ProjectModel)
    '    Get
    '        Return _CurrentProjects
    '    End Get
    '    Set(value As ObservableCollection(Of ProjectModel))
    '        _CurrentProjects = value
    '    End Set
    'End Property


    Private _CurrentProjects As ObservableCollection(Of AppDirectoryModel.ProjectModel)
    Public Property CurrentProjects As ObservableCollection(Of AppDirectoryModel.ProjectModel)
        Get
            Return _CurrentProjects
        End Get
        Set(value As ObservableCollection(Of AppDirectoryModel.ProjectModel))
            _CurrentProjects = value
        End Set
    End Property

    Public Class ProjectModel
        Private _Name As String
        Public Property Name As String
            Get
                Return _Name
            End Get
            Set(value As String)
                _Name = value
            End Set
        End Property

        Private _Directory As String
        Public Property Directory As String
            Get
                Return _Directory
            End Get
            Set(value As String)
                _Directory = value
            End Set
        End Property
    End Class

    Public ReadOnly Property ProjectsFileName As String
        Get
            Return AppDirectoryName & "\Projects"
        End Get
    End Property
End Class
