Public Class ConnectionViewModel
    Inherits BaseViewModel2(Of ConnectionModel)

    Private _ServerName As String
    Public Property ServerName As String
        Get
            Return Me._ServerName
        End Get
        Set(value As String)
            Me._ServerName = value
        End Set
    End Property

    Private _DataBaseName As String
    Public Property DataBaseName As String
        Get
            Return Me._DataBaseName
        End Get
        Set(value As String)
            Me._DataBaseName = value
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
        End Set
    End Property

    Private _OtherProperty As String
    Private Property OtherProperty As String
        Get
            Return Me._OtherProperty
        End Get
        Set(value As String)
            Me._OtherProperty = value
        End Set
    End Property

    Private Sub _GetConnectionString()
        Me.ConnectionString _
            = $"Server={Me.ServerName};DataBase={Me.DataBaseName};{Me._OtherProperty};"
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
        Dim b As Boolean : b = True
        If String.IsNullOrEmpty(Me.ServerName) Then
            b = False
        End If
        If String.IsNullOrEmpty(Me.DataBaseName) Then
            b = False
        End If
        If String.IsNullOrEmpty(Me._OtherProperty) Then
            b = False
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
        'Dim mysql As MySql

        'Call mysql.AccessTest()
        'mysql = Model.AccessTest()
        'If mysql.ResultFlag Then
        '    Me._GetTestTableName()
        'End If

        'If mysql.ResultFlag Then
        '    'mvm = New MigraterViewModel(Me.Model, Me.ViewModel)
        '    'ViewModel.ChangeContext(ViewModel.MAIN_VIEW, mvm.GetType.Name, mvm)
        'End If

        'mysql = Nothing
    End Sub

    Private Function _ConnectionCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ConnectionCommandEnableFlag
    End Function

    Protected Overrides Sub ViewInitializing()
    End Sub

    Sub New(ByRef m As Model,
            ByRef vm As ViewModel,
            ByRef adm As AppDirectoryModel,
            ByRef pim As ProjectInfoModel)

        Dim ip As InitializingProxy
        ip = AddressOf ViewInitializing

        Dim ccep(0) As CheckCommandEnabledProxy
        Dim ccep2 As CheckCommandEnabledProxy
        ccep(0) = AddressOf Me._CheckConnectionCommandEnabled
        ccep2 = [Delegate].Combine(ccep)

        Call Initializing(m, vm, adm, pim, ip, ccep2)
    End Sub
End Class
