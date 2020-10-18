Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ConnectionViewModel
    'Inherits BaseViewModel2(Of DBTestModel)
    Inherits BaseViewModel2

    Private Const _SUCCESS_MESSAGE As String = "接続に成功しました。"
    Private Const _NOCONNECTION As String = "まだ接続が成功していません。接続を定義してください"

    Public Overrides ReadOnly Property ViewType As String
        Get
            Return ViewModel.MAIN_VIEW
        End Get
    End Property

    Private _ServerName As String
    Public Property ServerName As String
        Get
            Return Me._ServerName
        End Get
        Set(value As String)
            Me._ServerName = value
            Model.Data.ServerName = value
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
            Model.Data.DataBaseName = value
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
            Model.Data.ConnectionString = value
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
            Call _CheckConnectionCommandEnabled()
        End Set
    End Property

    Private _InitializeMessage As String
    Public Property InitializeMessage As String
        Get
            Return Me._InitializeMessage
        End Get
        Set(value As String)
            Me._InitializeMessage = value
            RaisePropertyChanged("InitializeMessage")
        End Set
    End Property

    Private _ConnectionMessage As String
    Private Property ConnectionMessage As String
        Get
            Return Me._ConnectionMessage
        End Get
        Set(value As String)
            Me._ConnectionMessage = value
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
        Dim i As Integer : i = -1
        Dim a = Sub()
                    Call Model.DataSave(ProjectInfo)
                End Sub
        Dim sh As New SqlHandler With {
            .ConnectionString = Me.ConnectionString
        }
        Dim da As New DelegateAction With {
            .CanExecuteHandler = AddressOf sh.AccessTest,
            .ExecuteHandler = a
        }
        Select Case da.ExecuteIfCan(Nothing)
            Case 0
                Me.ConnectionMessage = _SUCCESS_MESSAGE
                Me.InitializeMessage = vbNullString
            Case 1000
                Me.ConnectionMessage = sh.Message
            Case Else
        End Select
    End Sub

    Private Function _ConnectionCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ConnectionCommandEnableFlag
    End Function

    Private Sub _ViewInitialize()
        Dim i = 0
        Dim b = False
        Dim sh As SqlHandler
        If Model.Data.ConnectionString <> Nothing Then
            Me.ConnectionString = Model.Data.ConnectionString
            i += 1
        End If
        If Model.Data.ServerName <> Nothing Then
            Me.ServerName = Model.Data.ServerName
            i += 1
        End If
        If Model.Data.DataBaseName <> Nothing Then
            Me.DataBaseName = Model.Data.DataBaseName
            i += 1
        End If
        If i >= 3 Then
            sh = New SqlHandler With {
                .ConnectionString = Me.ConnectionString
            }
            b = sh.AccessTest()
        End If
        If b Then
            Me.InitializeMessage = vbNullString
        Else
            Me.InitializeMessage = _NOCONNECTION
        End If
    End Sub

    Public Sub Initialize(ByRef m As Model,
                          ByRef vm As ViewModel,
                          ByRef adm As AppDirectoryModel,
                          ByRef pim As ProjectInfoModel)

        InitializeHandler = AddressOf _ViewInitialize
        CheckCommandEnabledHandler = [Delegate].Combine(
            New Action(AddressOf _CheckConnectionCommandEnabled)
        )

        Call BaseInitialize(m, vm, adm, pim)
    End Sub
End Class
