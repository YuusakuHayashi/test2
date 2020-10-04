Imports System.Collections.ObjectModel
Imports System.IO

Public Class AppDirectoryModel : Inherits ModelLoader(Of AppDirectoryModel)

    ' この静的メンバはアプリケーションディレクトリを表します
    Public Shared AppDirectoryName _
        = System.Environment.GetEnvironmentVariable("USERPROFILE") & "\hys"

    ' この静的メンバはアプリケーションディレクトリを表します
    Public Shared ModelFileName _
        = AppDirectoryModel.AppDirectoryName & "\Model.json"

    ' この静的メンバはアプリケーションディレクトリを表します
    Public Shared AppIniFileName _
        = AppDirectoryModel.AppDirectoryName & "\App.ini"

    Public Const DB_TEST As String = "データベーステスト(.dbt)"
    Public Shared ProjectKindList As List(Of String) _
        = New List(Of String) From {
            DB_TEST
        }

    Public Delegate Sub AppLaunchProxy()

    Private Sub CreateAppDirectory()
        Try
            Directory.CreateDirectory(AppDirectoryModel.AppDirectoryName)
        Catch ex As Exception
            Throw New Exception
        End Try
    End Sub

    Private Sub CreateAppModelFile()
        Try
            File.Create(AppDirectoryModel.ModelFileName)
        Catch ex As Exception
            Throw New Exception
        End Try
    End Sub

    Private Sub CreateAppIniFile()
        Try
            File.Create(AppDirectoryModel.AppIniFileName)
        Catch ex As Exception
            Throw New Exception
        End Try
    End Sub

    ' このプロパティはアプリケーションが持つプロジェクトの一覧を表します
    Private _CurrentProjects As ObservableCollection(Of ProjectInfoModel)
    Public Property CurrentProjects As ObservableCollection(Of ProjectInfoModel)
        Get
            Return _CurrentProjects
        End Get
        Set(value As ObservableCollection(Of ProjectInfoModel))
            _CurrentProjects = value
        End Set
    End Property

    ' この関数はアプリケーションディレクトリの存在チェックを行います
    Public Function CheckAppDirectory() As Integer
        Dim code As Integer : code = 99
        If Directory.Exists(AppDirectoryModel.AppDirectoryName) Then
            code = 10
            If File.Exists(AppDirectoryModel.AppIniFileName) Then
                code = 0
            End If
        End If
        CheckAppDirectory = code
    End Function

    ' この関数はアプリケーションの作成を行い、その結果を返します
    Public Function AppLaunch() As Integer
        Dim proxy(2) As AppLaunchProxy
        Dim proxy2 As AppLaunchProxy

        proxy(0) = AddressOf CreateAppDirectory
        proxy(1) = AddressOf CreateAppIniFile
        proxy(2) = AddressOf CreateAppModelFile

        Dim code As Integer : code = CheckAppDirectory()
        Try
            Select Case code
                Case 0
                Case 10
                Case 99
                    proxy2 = [Delegate].Combine(proxy)
                Case Else
            End Select
            If proxy2 IsNot Nothing Then
                If proxy2.GetInvocationList IsNot Nothing Then
                    Call proxy2()
                    code = 0
                End If
            End If
        Catch ex As Exception
        Finally
            AppLaunch = code
        End Try
    End Function

    Sub New()
        Dim proxy(0) As AppLaunchProxy
        Dim proxy2 As AppLaunchProxy

        Dim code As Integer : code = Me.CheckAppDirectory()

        proxy(0) = AddressOf Me.AppLaunch

        Select Case code
            Case 0
            Case 10
            Case 99
                proxy2 = [Delegate].Combine(proxy)
        End Select

        If proxy2 IsNot Nothing Then
            If proxy2.GetInvocationList IsNot Nothing Then
                Call proxy2()
            End If
        End If
    End Sub

    ' このプロパティはプロジェクトの構成を表します    
    'Public Class ProjectModel
    '    Private _Name As String
    '    Public Property Name As String
    '        Get
    '            Return _Name
    '        End Get
    '        Set(value As String)
    '            _Name = value
    '        End Set
    '    End Property

    '    Private _Kind As String
    '    Public Property Kind As String
    '        Get
    '            Return _Kind
    '        End Get
    '        Set(value As String)
    '            _Kind = value
    '        End Set
    '    End Property

    '    Private _Directory As String
    '    Public Property Directory As String
    '        Get
    '            Return _Directory
    '        End Get
    '        Set(value As String)
    '            _Directory = value
    '        End Set
    '    End Property

    '    ' プロジェクトのビューモデルファイルです
    '    Public ReadOnly Property ViewModelFileName As String
    '        Get
    '            Return Directory & "\ViewModel.json"
    '        End Get
    '    End Property

    '    ' プロジェクトのモデルファイルです
    '    Public ReadOnly Property ModelFileName As String
    '        Get
    '            Return Directory & "\Model.json"
    '        End Get
    '    End Property

    '    ' プロジェクトがアプリケーションにより作成されたことを表すファイルです
    '    Public ReadOnly Property IniFileName As String
    '        Get
    '            Return Directory & "\Project.ini"
    '        End Get
    '    End Property
    'End Class


    'Private Delegate Sub ProjectLaunchProxy(ByVal pm As AppDirectoryModel.ProjectModel)

    'Private Sub CreateProjectDirectory(ByVal pm As AppDirectoryModel.ProjectModel)
    '    Try
    '        Directory.CreateDirectory(pm.Directory)
    '    Catch ex As Exception
    '        Throw New Exception
    '    End Try
    'End Sub

    'Private Sub CreateProjectModelFile(ByVal pm As AppDirectoryModel.ProjectModel)
    '    Try
    '        File.Create(pm.ModelFileName)
    '    Catch ex As Exception
    '        Throw New Exception
    '    End Try
    'End Sub

    'Private Sub CreateProjectViewModelFile(ByVal pm As AppDirectoryModel.ProjectModel)
    '    Try
    '        File.Create(pm.ViewModelFileName)
    '    Catch ex As Exception
    '        Throw New Exception
    '    End Try
    'End Sub

    'Private Sub CreateProjectInitFile(ByVal pm As AppDirectoryModel.ProjectModel)
    '    Try
    '        File.Create(pm.IniFileName)
    '    Catch ex As Exception
    '        Throw New Exception
    '    End Try
    'End Sub

    '' この関数はプロジェクトの作成を行い、その結果を返します
    'Public Function ProjectLaunch(ByVal pm As AppDirectoryModel.ProjectModel) As Integer
    '    Dim proxy(3) As ProjectLaunchProxy
    '    Dim proxy2 As ProjectLaunchProxy

    '    proxy(0) = AddressOf CreateProjectDirectory
    '    proxy(1) = AddressOf CreateProjectInitFile
    '    proxy(2) = AddressOf CreateProjectModelFile
    '    proxy(3) = AddressOf CreateProjectViewModelFile

    '    Dim code As Integer : code = 99
    '    Try
    '        If Directory.Exists(pm.Directory) Then
    '            If File.Exists(pm.IniFileName) Then
    '            Else
    '                ' Nothing To Do
    '            End If
    '        Else
    '            proxy2 = [Delegate].Combine(proxy)
    '        End If

    '        If proxy2 IsNot Nothing Then
    '            If proxy2.GetInvocationList IsNot Nothing Then
    '                Call proxy2(pm)
    '                code = 0
    '            End If
    '        End If
    '    Catch ex As Exception
    '    Finally
    '        ProjectLaunch = code
    '    End Try
    'End Function
End Class
