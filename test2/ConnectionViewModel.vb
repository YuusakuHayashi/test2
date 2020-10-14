﻿Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ConnectionViewModel
    Inherits BaseViewModel2(Of DBTestModel)

    Private Const _SUCCESS_MESSAGE As String = "接続に成功しました"

    Private _ServerName As String
    Public Property ServerName As String
        Get
            Return Me._ServerName
        End Get
        Set(value As String)
            Me._ServerName = value
            Data.ServerName = value
            RaisePropertyChanged("ServerName")
            Call _GetConnectionString()
            Call _CheckConnectionCommandEnabled()
        End Set
    End Property

    Private _DataBaseName As String
    Public Property DataBaseName As String
        Get
            Return Me._DataBaseName
        End Get
        Set(value As String)
            Me._DataBaseName = value
            Data.DataBaseName = value
            RaisePropertyChanged("DataBaseName")
            Call _GetConnectionString()
            Call _CheckConnectionCommandEnabled()
        End Set
    End Property

    'Private _TestTableName As String
    'Public Property TestTableName As String
    '    Get
    '        Return Me._TestTableName
    '    End Get
    '    Set(value As String)
    '        Me._TestTableName = value
    '    End Set
    'End Property

    Private _ConnectionString As String
    Public Property ConnectionString As String
        Get
            Return Me._ConnectionString
        End Get
        Set(value As String)
            Me._ConnectionString = value
            Data.ConnectionString = value
            RaisePropertyChanged("ConnectionString")
            Call _GetOtherProperty()
        End Set
    End Property

    Private __OtherProperty As String
    Private Property _OtherProperty As String
        Get
            Return Me.__OtherProperty
        End Get
        Set(value As String)
            Me.__OtherProperty = value
            Data.OtherProperty = value
            Call _CheckConnectionCommandEnabled()
        End Set
    End Property

    Private _Message As String
    Private Property Message As String
        Get
            Return Me._Message
        End Get
        Set(value As String)
            Me._Message = value
        End Set
    End Property

    Private Sub _GetConnectionString()
        Me.ConnectionString _
            = $"Server={Me.ServerName};DataBase={Me.DataBaseName};{Me._OtherProperty};"
    End Sub

    ' 
    Private Sub _GetOtherProperty()
        Dim txt As String
        txt = Me.ConnectionString
        If Not String.IsNullOrEmpty(Me.ConnectionString) Then
            If txt.Contains($"Server={Me.ServerName};DataBase={Me.DataBaseName};") Then
                txt = Me.ConnectionString.Replace($"Server={Me.ServerName};DataBase={Me.DataBaseName};", "")
                If txt.Contains(";") Then
                    txt = txt.TrimEnd(";")
                End If
            End If
        End If
        Me._OtherProperty = txt
    End Sub


    ' コマンドプロパティ（接続確認）
    Private _ConnectionCommand As ICommand
    Public ReadOnly Property ConnectionCommand As ICommand
        Get
            If Me._ConnectionCommand Is Nothing Then
                Me._ConnectionCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ConnectionCommandExecute,
                    .CanExecuteHandler = AddressOf _ConnectionCommandCanExecute
                }
                Return Me._ConnectionCommand
            Else
                Return Me._ConnectionCommand
            End If
        End Get
    End Property

    'コマンド実行可否のチェック（接続確認）
    Private Sub _CheckConnectionCommandEnabled()
        Dim b As Boolean : b = False
        If Not String.IsNullOrEmpty(Me.ServerName) Then
            If Not String.IsNullOrEmpty(Me.DataBaseName) Then
                If Not String.IsNullOrEmpty(Me._OtherProperty) Then
                    b = True
                End If
            End If
        End If
        Me._ConnectionCommandEnableFlag = b
    End Sub

    ' コマンド実行可否のフラグ（接続確認）
    Private __ConnectionCommandEnableFlag As Boolean
    Public Property _ConnectionCommandEnableFlag As Boolean
        Get
            Return Me.__ConnectionCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ConnectionCommandEnableFlag = value
            RaisePropertyChanged("_ConnectionCommandEnableFlag")
            CType(ConnectionCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    ' コマンド実行（接続確認）
    Private Sub _ConnectionCommandExecute(ByVal parameter As Object)
        Dim sh As New SqlHandler With {
            .ConnectionString = Me.ConnectionString
        }
        If sh.AccessTest() Then
            Me.Message = _SUCCESS_MESSAGE
        Else
            Me.Message = sh.Message
        End If
    End Sub

    Private Function _ConnectionCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ConnectionCommandEnableFlag
    End Function

    Public Sub MyInitializing(ByRef m As Model,
                              ByRef vm As ViewModel,
                              ByRef adm As AppDirectoryModel,
                              ByRef pim As ProjectInfoModel)

        CheckCommandEnabledHandler = [Delegate].Combine(
            New Action(AddressOf _CheckConnectionCommandEnabled)
        )

        Call Initializing(m, vm, adm, pim)
    End Sub
End Class
