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

    Public Sub AppSave()
        Call Me.ModelSave(
            AppDirectoryModel.ModelFileName, Me
        )
    End Sub

    Public Const DBTEST As String = "データベーステスト(.dbt)"
    Public Const DBTEST_IMAGE_URL As String = "\\Coral\個人情報-林祐\project\wpf\test2\test2\rpa.ico"

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
                url = DBTEST_IMAGE_URL
            Case Else
                url = DBTEST_IMAGE_URL
        End Select
        AssignImageOfProject = url
    End Function

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
        Dim proxy As Action

        Dim code As Integer : code = CheckAppDirectory()
        Try
            Select Case code
                Case 0
                Case 10
                    ' 失敗用のロジックが必要
                Case 99
                    proxy = [Delegate].Combine(
                        New Action(AddressOf CreateAppDirectory),
                        New Action(AddressOf CreateAppIniFile),
                        New Action(AddressOf CreateAppModelFile)
                    )
                Case Else
            End Select
            If proxy IsNot Nothing Then
                Call proxy()
                code = 0
            End If
        Catch ex As Exception
        Finally
            AppLaunch = code
        End Try
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
        Dim proxy As Action
        Dim code As Integer : code = Me.CheckAppDirectory()

        Select Case code
            Case 0
            Case 10
            Case 99
                proxy = [Delegate].Combine(
                    New Action(AddressOf AppLaunch)
                )
        End Select

        If proxy IsNot Nothing Then
            Call proxy()
        End If
    End Sub

    Public Sub Initialize()
        Dim bi As BitmapImage
        For Each p In Me.CurrentProjects
            If String.IsNullOrEmpty(p.ImageFileName) Then
                p.ImageFileName = DBTEST_IMAGE_URL
            End If
            bi = New BitmapImage
            bi.BeginInit()
            bi.UriSource = New Uri((p.ImageFileName), UriKind.Absolute)
            bi.EndInit()
            p.Image = bi
        Next
    End Sub
End Class
