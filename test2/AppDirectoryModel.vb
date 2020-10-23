Imports System.Collections.ObjectModel
Imports System.IO

Public Class AppDirectoryModel : Inherits JsonHandler(Of AppDirectoryModel)
'Public Class AppDirectoryModel
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

    Public Sub AppSave()
        Call Me.ModelSave(
            AppDirectoryModel.ModelFileName, Me
        )
    End Sub

    Private Shared _DBTEST_IMAGE As String _
        = _AppImageDirectory & "\rpa.ico"

    Public Const DBTEST As String = "データベーステスト(.dbt)"

    Public Shared ProjectKindList As List(Of String) _
        = New List(Of String) From {
            DBTEST
        }

    Private Sub _AssignProjectIndex(ByRef project As ProjectInfoModel)
        Dim idx = 1
        For Each p In Me.CurrentProjects
            If p.[Index] = idx Then
                idx += 1
            End If
        Next
        project.[Index] = idx
    End Sub

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

    Public Function AssignImageOfProject(ByVal kind As String)
        Dim url = vbNullString
        Select Case kind
            Case DBTEST
                url = _DBTEST_IMAGE
            Case Else
                url = _DBTEST_IMAGE
        End Select
        AssignImageOfProject = url
    End Function

    Private Sub _Make(ByVal fds As List(Of Object))
        Dim fi As FileInfo
        Dim di As DirectoryInfo
        For Each fd In fds
            fd.Create()
            'If TryCast(fd, FileInfo) IsNot Nothing Then
            '    fi = CType(fd, FileInfo)
            '    fi.Create()
            '    Continue For
            'End If
            'If TryCast(fd, DirectoryInfo) IsNot Nothing Then
            '    di = CType(fd, DirectoryInfo)
            '    di.Create()
            '    Continue For
            'End If
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

    Public Sub PushProject(ByVal project As ProjectInfoModel)
        Dim [new] As New ObservableCollection(Of ProjectInfoModel)

        [new].Add(project)
        If project.[Index] = 0 Then
            Call _AssignProjectIndex(project)
        End If

        For Each p In Me.CurrentProjects
            If project.[Index] <> p.[Index] Then
                [new].Add(p)
            End If
            If [new].Count >= 5 Then
                Exit For
            End If
        Next

        Me.CurrentProjects = [new]
    End Sub

    Sub New()
        Call AppLaunch()
    End Sub

    Public Sub Initialize()
        Dim bi As BitmapImage
        For Each p In Me.CurrentProjects
            If String.IsNullOrEmpty(p.ImageFileName) Then
                p.ImageFileName = _DBTEST_IMAGE
            End If
            bi = New BitmapImage
            bi.BeginInit()
            bi.UriSource = New Uri((p.ImageFileName), UriKind.Absolute)
            bi.EndInit()
            p.Image = bi
        Next
    End Sub
End Class
