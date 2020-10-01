'Imports System.Data

Public Class InitViewModel
    '    Inherits ProjectBaseViewModel(Of InitViewModel)

    '    ' サーバ名
    '    ' 接続文字列にバインド
    '    ' コマンド実行フラグをチェック
    '    Private _ServerName As String
    '    Public Property ServerName As String
    '        Get
    '            Return Me._ServerName
    '        End Get
    '        Set(value As String)
    '            Me._ServerName = value
    '            RaisePropertyChanged("ServerName")
    '            Call Me._GetConnectionString()
    '            Call Me._CheckInitCommandEnabled()
    '        End Set
    '    End Property


    '    ' ＤＢ名
    '    ' 変更を接続文字列にバインドさせる
    '    ' コマンド実行フラグをチェック
    '    Private _DataBaseName As String
    '    Public Property DataBaseName As String
    '        Get
    '            Return Me._DataBaseName
    '        End Get
    '        Set(value As String)
    '            Me._DataBaseName = value
    '            RaisePropertyChanged("DataBaseName")
    '            Call Me._GetConnectionString()
    '            Call Me._CheckInitCommandEnabled()
    '        End Set
    '    End Property


    '    ' コマンド実行フラグをチェック
    '    Private _ConnectionString As String
    '    Public Property ConnectionString As String
    '        Get
    '            Return Me._ConnectionString
    '        End Get
    '        Set(value As String)
    '            Me._ConnectionString = value
    '            RaisePropertyChanged("ConnectionString")
    '            Call Me._GetOtherProperty()
    '            Call Me._CheckInitCommandEnabled()
    '        End Set
    '    End Property


    '    Private __OtherProperty As String
    '    Private Property _OtherProperty As String
    '        Get
    '            Return Me.__OtherProperty
    '        End Get
    '        Set(value As String)
    '            Me.__OtherProperty = value
    '        End Set
    '    End Property


    '    Private Sub _GetOtherProperty()
    '        Dim txt As String
    '        txt = Me.ConnectionString
    '        If String.IsNullOrEmpty(Me.ConnectionString) Then
    '            '
    '        Else
    '            If txt.Contains($"Server={Me.ServerName};DataBase={Me.DataBaseName};") Then
    '                txt = Me.ConnectionString.Replace($"Server={Me.ServerName};DataBase={Me.DataBaseName};", "")
    '                If txt.Contains(";") Then
    '                    txt = txt.TrimEnd(";")
    '                End If
    '            End If
    '        End If
    '        Me._OtherProperty = txt
    '    End Sub


    '    Private Sub _GetConnectionString()
    '        Me.ConnectionString = $"Server={Me.ServerName};DataBase={Me.DataBaseName};{Me._OtherProperty};"
    '    End Sub


    '    Public ReadOnly Property ServerNameError As String
    '        Get
    '            Return vbNullString
    '        End Get
    '    End Property

    '    Public ReadOnly Property DataBaseNameError As String
    '        Get
    '            Return vbNullString
    '        End Get
    '    End Property

    '    Public ReadOnly Property ConnectionStringError As String
    '        Get
    '            Return vbNullString
    '        End Get
    '    End Property


    '    Private _InitCommand As ICommand
    '    Public ReadOnly Property InitCommand As ICommand
    '        Get
    '            If Me._InitCommand Is Nothing Then
    '                Me._InitCommand = New DelegateCommand With {
    '                    .ExecuteHandler = AddressOf _InitCommandExecute,
    '                    .CanExecuteHandler = AddressOf _InitCommandCanExecute
    '                }
    '                Return Me._InitCommand
    '            Else
    '                Return Me._InitCommand
    '            End If
    '        End Get
    '    End Property


    '    'コマンド実行可否のチェック
    '    Private Sub _CheckInitCommandEnabled()
    '        Dim b As Boolean : b = True
    '        If Me.ServerName = vbNullString Then
    '            b = False
    '        End If
    '        If Me.DataBaseName = vbNullString Then
    '            b = False
    '        End If
    '        If Me._OtherProperty = vbNullString Then
    '            b = False
    '        End If
    '        If Me.ConnectionString = vbNullString Then
    '            b = False
    '        End If
    '        If Me.ServerNameError <> vbNullString Then
    '            b = False
    '        End If
    '        If Me.DataBaseNameError <> vbNullString Then
    '            b = False
    '        End If
    '        If Me.ConnectionStringError <> vbNullString Then
    '            b = False
    '        End If
    '        Me._InitEnableFlag = b
    '    End Sub


    '    'コマンド実行可否のフラグ
    '    Private __InitEnableFlag As Boolean
    '    Public Property _InitEnableFlag As Boolean
    '        Get
    '            Return Me.__InitEnableFlag
    '        End Get
    '        Set(value As Boolean)
    '            Me.__InitEnableFlag = value
    '            RaisePropertyChanged("_InitEnableFlag")
    '            CType(InitCommand, DelegateCommand).RaiseCanExecuteChanged()
    '        End Set
    '    End Property


    '    Private Sub _InitCommandExecute(ByVal parameter As Object)
    '        Dim mysql As MySql
    '        Dim mvm As MigraterViewModel

    '        Me.Model.ServerName = Me.ServerName
    '        Me.Model.DataBaseName = Me.DataBaseName
    '        Me.Model.ConnectionString = Me.ConnectionString
    '        Me.Model.OtherProperty = Me._OtherProperty

    '        'Call mysql.AccessTest()
    '        mysql = Model.AccessTest()
    '        If mysql.ResultFlag Then
    '            Me._GetTestTableName()
    '        End If

    '        If mysql.ResultFlag Then
    '            'mvm = New MigraterViewModel(Me.Model, Me.ViewModel)
    '            'ViewModel.ChangeContext(ViewModel.MAIN_VIEW, mvm.GetType.Name, mvm)
    '        End If

    '        mysql = Nothing
    '    End Sub

    '    Private Sub _GetTestTableName()
    '        Dim mysql As MySql
    '        Dim i As Integer
    '        Dim pseudonym As String
    '        Dim exit_flg As Boolean
    '        Dim dt As DataTable
    '        Dim list As New List(Of String)

    '        mysql = Model.GetUserTables()
    '        dt = mysql.Result.Tables(0)

    '        For Each r In dt.Rows
    '            list.Add(r("name"))
    '        Next


    '        i = 0
    '        Do
    '            pseudonym = "TEST" & (i).ToString
    '            exit_flg = False
    '            For Each l In list
    '                If l = pseudonym Then
    '                    exit_flg = False
    '                    Exit For
    '                End If
    '            Next

    '            If exit_flg Then
    '                Exit Do
    '            End If
    '        Loop Until 1 = 1

    '        Model.TestTableName = pseudonym
    '    End Sub

    '    Private Function _InitCommandCanExecute(ByVal parameter As Object) As Boolean
    '        Return Me._InitEnableFlag
    '    End Function

    '    Protected Overrides Sub MyInitializing()
    '        Me.ServerName = Me.Model.ServerName
    '        Me.DataBaseName = Me.Model.DataBaseName
    '        Me._OtherProperty = Me.Model.OtherProperty
    '        Call Me._GetConnectionString()
    '    End Sub

    '    Protected Overrides Sub ContextModelCheck()
    '        ViewModel.SetContext(ViewModel.MAIN_VIEW, Me.GetType.Name, Me)
    '    End Sub

    '    Sub New(ByRef m As Model,
    '            ByRef vm As ViewModel)

    '        Dim ccep(0) As CheckCommandEnabledProxy
    '        Dim ccep2 As CheckCommandEnabledProxy

    '        ccep(0) = AddressOf Me._CheckInitCommandEnabled
    '        ccep2 = [Delegate].Combine(ccep)


    '        ' ビューモデルの設定
    '        Initializing(m, vm, ccep2)
    '    End Sub
End Class
