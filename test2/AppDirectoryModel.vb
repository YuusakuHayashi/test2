Imports System.Collections.ObjectModel
Imports System.IO

Public Class AppDirectoryModel

    ' このプロパティはアプリケーションディレクトリを表します
    Public ReadOnly Property AppDirectoryName As String
        Get
            Return System.Environment.GetEnvironmentVariable("USERPROFILE") & "\hys"
        End Get
    End Property

    ' このプロパティはアプリケーションが持つプロジェクトの一覧を表します
    Private _CurrentProjects As ObservableCollection(Of AppDirectoryModel.ProjectModel)
    Public Property CurrentProjects As ObservableCollection(Of AppDirectoryModel.ProjectModel)
        Get
            Return _CurrentProjects
        End Get
        Set(value As ObservableCollection(Of AppDirectoryModel.ProjectModel))
            _CurrentProjects = value
        End Set
    End Property

    ' このプロパティはプロジェクトの構成を表します    
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

        Private _Kind As String
        Public Property Kind As String
            Get
                Return _Kind
            End Get
            Set(value As String)
                _Kind = value
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

        Public ReadOnly Property IniFileName As String
            Get
                Return Directory & "\Project.ini"
            End Get
        End Property
    End Class

    Private Delegate Sub ProjectLaunchProxy(ByVal pm As AppDirectoryModel.ProjectModel)

    Private Sub CreateProjectDirectory(ByVal pm As AppDirectoryModel.ProjectModel)
        Try
            Directory.CreateDirectory(pm.Directory)
        Catch ex As Exception
            Throw New Exception
        End Try
    End Sub

    Private Sub CreateProjectInitFile(ByVal pm As AppDirectoryModel.ProjectModel)
        Try
            File.Create(pm.IniFileName)
        Catch ex As Exception
            Throw New Exception
        End Try
    End Sub

    ' この関数はプロジェクトの作成を行い、その結果を返します
    Public Function ProjectLaunch(ByVal pm As AppDirectoryModel.ProjectModel) As Integer
        Dim proxy(1) As ProjectLaunchProxy
        Dim proxy2 As ProjectLaunchProxy

        proxy(0) = AddressOf CreateProjectDirectory
        proxy(1) = AddressOf CreateProjectInitFile

        Dim code As Integer : code = 99
        Try
            If Directory.Exists(pm.Directory) Then
                If File.Exists(pm.IniFileName) Then
                Else
                    ' Nothing To Do
                End If
            Else
                proxy2 = [Delegate].Combine(proxy)
            End If
            If proxy2 IsNot Nothing Then
                If proxy2.GetInvocationList IsNot Nothing Then
                    Call proxy2(pm)
                    code = 0
                End If
            End If
        Catch ex As Exception
        Finally
            ProjectLaunch = code
        End Try
    End Function

    Public ReadOnly Property ProjectsFileName As String
        Get
            Return AppDirectoryName & "\Projects"
        End Get
    End Property
End Class
