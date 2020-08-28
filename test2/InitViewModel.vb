Public Class InitViewModel : Inherits ViewModel
    ' モデル
    Private __Model As Model
    Private Property _Model As Model
        Get
            Return Me.__Model
        End Get
        Set(value As Model)
            Me.__Model = value
            RaisePropertyChanged("Model")
        End Set
    End Property


    ' 初期画面用フラグ
    Private _NextFlag As Boolean
    Public Property NextFlag As Boolean
        Get
            Return Me._NextFlag
        End Get
        Set(value As Boolean)
            Me._NextFlag = value
            RaisePropertyChanged("NextFlag")
        End Set
    End Property


    ' サーバ名
    ' 接続文字列にバインド
    ' コマンド実行フラグをチェック
    Private _ServerName As String
    Public Property ServerName As String
        Get
            Return Me._ServerName
        End Get
        Set(value As String)
            Me._ServerName = value
            RaisePropertyChanged("ServerName")
            Call Me._GetConnectionString()
            Call Me._CheckInitEnableFlag()
        End Set
    End Property


    ' ＤＢ名
    ' 変更を接続文字列にバインドさせる
    ' コマンド実行フラグをチェック
    Private _DataBaseName As String
    Public Property DataBaseName As String
        Get
            Return Me._DataBaseName
        End Get
        Set(value As String)
            Me._DataBaseName = value
            RaisePropertyChanged("DataBaseName")
            Call Me._GetConnectionString()
            Call Me._CheckInitEnableFlag()
        End Set
    End Property


    ' コマンド実行フラグをチェック
    Private _ConnectionString As String
    Public Property ConnectionString As String
        Get
            Return Me._ConnectionString
        End Get
        Set(value As String)
            Me._ConnectionString = value
            RaisePropertyChanged("ConnectionString")
            Call Me._GetOtherProperty()
            Call Me._CheckInitEnableFlag()
        End Set
    End Property


    Private __OtherProperty As String
    Private Property _OtherProperty As String
        Get
            Return Me.__OtherProperty
        End Get
        Set(value As String)
            Me.__OtherProperty = value
        End Set
    End Property


    Private Sub _GetOtherProperty()
        Dim txt As String
        txt = Me.ConnectionString
        If String.IsNullOrEmpty(Me.ConnectionString) Then
            '
        Else
            If txt.Contains($"Server={Me.ServerName};DataBase={Me.DataBaseName};") Then
                txt = Me.ConnectionString.Replace($"Server={Me.ServerName};DataBase={Me.DataBaseName};", "")
                If txt.Contains(";") Then
                    txt = txt.TrimEnd(";")
                End If
            End If
        End If
        Me._OtherProperty = txt
    End Sub


    Private Sub _GetConnectionString()
        Me.ConnectionString = $"Server={Me.ServerName};DataBase={Me.DataBaseName};{Me._OtherProperty};"
    End Sub


    Public ReadOnly Property ServerNameError As String
        Get
            Return vbNullString
        End Get
    End Property

    Public ReadOnly Property DataBaseNameError As String
        Get
            Return vbNullString
        End Get
    End Property

    Public ReadOnly Property ConnectionStringError As String
        Get
            Return vbNullString
        End Get
    End Property


    Private _InitCommand As ICommand
    Public ReadOnly Property InitCommand As ICommand
        Get
            If Me._InitCommand Is Nothing Then
                Me._InitCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _InitCommandExecute,
                    .CanExecuteHandler = AddressOf _InitCommandCanExecute
                }
                Return Me._InitCommand
            Else
                Return Me._InitCommand
            End If
        End Get
    End Property


    'コマンド実行可否のチェック
    Private Sub _CheckInitEnableFlag()
        Dim b As Boolean : b = True
        If Me.ServerName = vbNullString Then
            b = False
        End If
        If Me.DataBaseName = vbNullString Then
            b = False
        End If
        If Me._OtherProperty = vbNullString Then
            b = False
        End If
        If Me.ConnectionString = vbNullString Then
            b = False
        End If
        If Me.ServerNameError <> vbNullString Then
            b = False
        End If
        If Me.DataBaseNameError <> vbNullString Then
            b = False
        End If
        If Me.ConnectionStringError <> vbNullString Then
            b = False
        End If
        Me._InitEnableFlag = b
    End Sub


    'コマンド実行可否のフラグ
    Private __InitEnableFlag As Boolean
    Public Property _InitEnableFlag As Boolean
        Get
            Return Me.__InitEnableFlag
        End Get
        Set(value As Boolean)
            Me.__InitEnableFlag = value
            RaisePropertyChanged("InitEnableFlag")
            CType(InitCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property


    Private Sub _InitCommandExecute(ByVal parameter As Object)
        Dim pm As New ProjectModel
        Dim proxy As ProjectCheckProxy

        Me._Model.ServerName = Me.ServerName
        Me._Model.DataBaseName = Me.DataBaseName
        Me._Model.ConnectionString = Me.ConnectionString
        Me._Model.OtherProperty = Me._OtherProperty

        Call Me._Model.AccessTest()

        If Me._Model.AccessResult Then
            Me._Model.ChangePageStrings = {Me.GetType.Name, "MenuViewModel"}
        End If
    End Sub


    Private Function _InitCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._InitEnableFlag
    End Function


    Private Sub _DefaultSet()
        Me.ServerName = vbNullString
        Me.DataBaseName = vbNullString
        Me.ConnectionString = vbNullString
    End Sub



    Sub New(ByRef m As Model)
        If m IsNot Nothing Then
            ' 接続テスト
            Call m.AccessTest()
            Me.NextFlag = True
            If Not m.AccessResult Then
                ' 接続失敗
                If String.IsNullOrEmpty(m.ServerName) Then
                    m.ServerName = vbNullString
                End If
                If String.IsNullOrEmpty(m.DataBaseName) Then
                    m.DataBaseName = vbNullString
                End If
                If String.IsNullOrEmpty(m.OtherProperty) Then
                    m.OtherProperty = vbNullString
                End If
            End If
        Else
            ' モデルなし
            m = New Model
            m.ServerName = vbNullString
            m.DataBaseName = vbNullString
            m.OtherProperty = vbNullString
        End If

        Me._Model = m

        Me.ServerName = m.ServerName
        Me.DataBaseName = m.DataBaseName
        Me._OtherProperty = m.OtherProperty
        Call Me._GetConnectionString()
    End Sub
End Class
