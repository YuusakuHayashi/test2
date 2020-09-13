Public Class InitViewModel
    Inherits ProjectBaseViewModel(Of InitViewModel)

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
            Call Me._CheckInitCommandEnabled()
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
            Call Me._CheckInitCommandEnabled()
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
            Call Me._CheckInitCommandEnabled()
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
    Private Sub _CheckInitCommandEnabled()
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
            RaisePropertyChanged("_InitEnableFlag")
            CType(InitCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property


    Private Sub _InitCommandExecute(ByVal parameter As Object)
        Dim mvm As MigraterViewModel
        Me.Model.ServerName = Me.ServerName
        Me.Model.DataBaseName = Me.DataBaseName
        Me.Model.ConnectionString = Me.ConnectionString
        Me.Model.OtherProperty = Me._OtherProperty

        Call Me.Model.AccessTest()

        If Me.Model.AccessResult Then
            ModelSave(Of Model)(Me.Model.CurrentModelJson, Me.Model)
            Me.ViewModel.ContextModel = New MigraterViewModel(Me.Model, Me.ViewModel)
        End If

    End Sub


    Private Function _InitCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._InitEnableFlag
    End Function

    Protected Overrides Sub MyInitializing()
        Me.ServerName = Me.Model.ServerName
        Me.DataBaseName = Me.Model.DataBaseName
        Me._OtherProperty = Me.Model.OtherProperty
        Call Me._GetConnectionString()
    End Sub

    Protected Overrides Sub ContextModelCheck()
        Dim b As Boolean : b = False
        If String.IsNullOrEmpty(Me.ServerName) Then
            b = True
        End If
        If String.IsNullOrEmpty(Me.DataBaseName) Then
            b = True
        End If
        If String.IsNullOrEmpty(Me._OtherProperty) Then
            b = True
        End If
        If b Then
            ViewModel.ContextModel = Me
        End If
    End Sub

    Sub New(ByRef m As Model,
            ByRef vm As ViewModel)

        Dim ccep(0) As CheckCommandEnabledProxy
        Dim ccep2 As CheckCommandEnabledProxy

        ccep(0) = AddressOf Me._CheckInitCommandEnabled
        ccep2 = [Delegate].Combine(ccep)


        ' ビューモデルの設定
        Initializing(m, vm, ccep2)
    End Sub
End Class
