Imports System.Collections.ObjectModel
Imports System.IO
Imports Newtonsoft.Json

Public Class AppDirectoryModel : Inherits JsonHandler(Of AppDirectoryModel)
    ' この静的メンバはアプリケーションディレクトリを表します
    Public Shared AppDirectoryName _
        = System.Environment.GetEnvironmentVariable("USERPROFILE") & "\hys"

    Public Shared AppIniFileName _
        = AppDirectoryModel.AppDirectoryName & "\App.ini"

    ' この静的メンバはアプリケーションディレクトリを表します
    Public Shared ModelFileName _
        = AppDirectoryModel.AppDirectoryName & "\Model.json"

    Public Shared AppImageDirectory _
        = AppDirectoryModel.AppDirectoryName & "\Image"

    Public Const DBTEST As String = "データベーステスト"
    Public Const RpaProject As String = "RPAプロジェクト"

    Public Shared ProjectKindList As List(Of String) _
        = New List(Of String) From {
            DBTEST,
            RpaProject
        }

    Private Shared NoIcon _
        = AppDirectoryModel.AppImageDirectory & "\NoIcon.png"

    Private _ProjectInfo As ProjectInfoModel
    <JsonIgnore>
    Public Property ProjectInfo As ProjectInfoModel
        Get
            If Me._ProjectInfo Is Nothing Then
                Me._ProjectInfo = New ProjectInfoModel
            End If
            Return _ProjectInfo
        End Get
        Set(value As ProjectInfoModel)
            _ProjectInfo = value
        End Set
    End Property

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

    Private __IconDictionary As Dictionary(Of String, BitmapImage)
    Private Property _IconDictionary As Dictionary(Of String, BitmapImage)
        Get
            If Me.__IconDictionary Is Nothing Then
                Me.__IconDictionary = New Dictionary(Of String, BitmapImage)
            End If
            Return __IconDictionary
        End Get
        Set(value As Dictionary(Of String, BitmapImage))
            Me.__IconDictionary = value
        End Set
    End Property

    ' このプロパティはアプリケーションが持つプロジェクトの一覧を表します
    Private _CurrentProjects As ObservableCollection(Of ProjectInfoModel)
    Public Property CurrentProjects As ObservableCollection(Of ProjectInfoModel)
        Get
            If Me._CurrentProjects Is Nothing Then
                Me._CurrentProjects = New ObservableCollection(Of ProjectInfoModel)
            End If
            Return _CurrentProjects
        End Get
        Set(value As ObservableCollection(Of ProjectInfoModel))
            Me._CurrentProjects = value
            Call _AssignIconOfProject(value)
        End Set
    End Property

    Private _FixedProjects As ObservableCollection(Of ProjectInfoModel)
    Public Property FixedProjects As ObservableCollection(Of ProjectInfoModel)
        Get
            If Me._FixedProjects Is Nothing Then
                Me._FixedProjects = New ObservableCollection(Of ProjectInfoModel)
            End If
            Return _FixedProjects
        End Get
        Set(value As ObservableCollection(Of ProjectInfoModel))
            _FixedProjects = value
            Call _AssignIconOfProject(value)
        End Set
    End Property

    Private Sub _AssignIconOfProject(ByRef projects As ObservableCollection(Of ProjectInfoModel))
        Dim iconf As String
        For Each p In projects
            If p.Icon Is Nothing Then
                If Not String.IsNullOrEmpty(p.IconFileName) Then
                    iconf = IIf(File.Exists(p.IconFileName), p.IconFileName, NOICON)
                    If Not Me._IconDictionary.ContainsKey(iconf) Then
                        Call _RegisterIcon(iconf)
                    End If
                    p.Icon = Me._IconDictionary(iconf)
                End If
            End If
        Next
    End Sub

    Private Sub _RegisterIcon(ByVal f As String)
        Dim bi = New BitmapImage
        bi.BeginInit()
        bi.UriSource = New Uri((f), UriKind.Absolute)
        bi.EndInit()

        Me._IconDictionary(f) = bi
    End Sub

    Private Sub _Make(ByVal fds As List(Of Object))
        Dim fi As FileInfo
        Dim di As DirectoryInfo
        For Each fd In fds
            fd.Create()
        Next
    End Sub

    ' この関数はアプリケーションディレクトリの存在チェックを行います
    Public Function CheckAppDirectory() As Integer
        Dim code = 999
        If Directory.Exists(AppDirectoryModel.AppDirectoryName) Then
            code = 998
            If File.Exists(AppDirectoryModel.AppIniFileName) Then
                code = 997
                If File.Exists(AppDirectoryModel.ModelFileName) Then
                    code = 996
                    If Directory.Exists(AppDirectoryModel.AppImageDirectory) Then
                        code = 0
                    End If
                End If
            End If
        End If
        CheckAppDirectory = code
    End Function


    ' この関数はアプリケーションの作成を行い、その結果を返します
    Public Function AppLaunch() As Integer
        Dim fds As New List(Of Object)

        Dim code = CheckAppDirectory()
        If code = 998 Then
            MsgBox("Error AppLaunch = 998")
        End If
        If code >= 996 Then
            fds.Add(New DirectoryInfo(AppDirectoryModel.AppImageDirectory))
        End If
        If code >= 997 Then
            fds.Add(New FileInfo(AppDirectoryModel.ModelFileName))
        End If
        If code >= 999 Then
            fds.Add(New FileInfo(AppDirectoryModel.AppIniFileName))
            fds.Add(New DirectoryInfo(AppDirectoryModel.AppDirectoryName))
        End If

        If fds.Count > 0 Then
            Call _Make(fds)
        End If

        AppLaunch = code
    End Function

    Sub New()
        Call AppLaunch()
    End Sub
End Class
