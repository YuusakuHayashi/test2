Imports System.Collections.ObjectModel
Imports System.IO

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
                    ' 失敗用のロジックが必要
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
        proxy(0) = AddressOf Me.AppLaunch

        Dim code As Integer : code = Me.CheckAppDirectory()

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
End Class
