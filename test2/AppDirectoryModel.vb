Imports System.Collections.ObjectModel
Imports System.IO
Imports Newtonsoft.Json

Public Class AppDirectoryModel : Inherits JsonHandler(Of AppDirectoryModel)
    ' この静的メンバはアプリケーションディレクトリを表します
    Public Shared AppDirectoryName _
        = System.Environment.GetEnvironmentVariable("USERPROFILE") & "\hys"

    ' この静的メンバはアプリケーションディレクトリを表します
    Public Shared ModelFileName _
        = AppDirectoryModel.AppDirectoryName & "\Model.json"

    ' この静的メンバはアプリケーションディレクトリを表します
    Public Shared AppIniFileName _
        = AppDirectoryModel.AppDirectoryName & "\App.ini"

    Private Shared _AppImageDirectory _
        = AppDirectoryModel.AppDirectoryName & "\Image"

    Private Shared _DBTEST_IMAGE As String _
        = _AppImageDirectory & "\rpa.ico"

    Public Const DBTEST As String = "データベーステスト"
    Public Const RpaProject As String = "RPAプロジェクト"

    Public Shared ProjectKindList As List(Of String) _
        = New List(Of String) From {
            DBTEST,
            RpaProject
        }

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
            _CurrentProjects = value
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
        End Set
    End Property

    Public Function AssignIconOfProject(ByVal kind As String)
        Dim url = vbNullString
        Select Case kind
            Case DBTEST
                url = _DBTEST_IMAGE
            Case Else
                url = _DBTEST_IMAGE
        End Select
        AssignIconOfProject = url
    End Function

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
                    If Directory.Exists(AppDirectoryModel._AppImageDirectory) Then
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
            fds.Add(New DirectoryInfo(AppDirectoryModel._AppImageDirectory))
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

    Public Sub Initialize()
        Dim bi As BitmapImage
        For Each p In Me.CurrentProjects
            If String.IsNullOrEmpty(p.IconFileName) Then
                p.IconFileName = _DBTEST_IMAGE
            End If
            bi = New BitmapImage
            bi.BeginInit()
            bi.UriSource = New Uri((p.IconFileName), UriKind.Absolute)
            bi.EndInit()
            p.Icon = bi
        Next
        For Each p In Me.FixedProjects
            If String.IsNullOrEmpty(p.IconFileName) Then
                p.IconFileName = _DBTEST_IMAGE
            End If
            bi = New BitmapImage
            bi.BeginInit()
            bi.UriSource = New Uri((p.IconFileName), UriKind.Absolute)
            bi.EndInit()
            p.Icon = bi
        Next
    End Sub

    Sub New()
        Call AppLaunch()
    End Sub
End Class
