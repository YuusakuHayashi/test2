Public Class InitViewModel : Inherits ViewModel

    Private Const DEFAULT_SERVER_NAME As String = ""
    Private Const DEFAULT_DATABASE_NAME As String = ""
    Private Const DEFUALT_OTHERPROPERTY As String = "Trusted_Connection=True"
    Private Const DEFAULT_CONNECTIONSTRING As String = ""

    Private Const INITVIEW_FILE As String = "InitView.txt"


    'セーブ用にモデルを保持
    'ビューモデルをテキストにデシリアライズしようとすると
    'Newtonsoft.jsonでは、コマンドがデシリアライズ不可
    Private __Model As InitModel
    Private Property _Model As InitModel
        Get
            Return Me.__Model
        End Get
        Set(value As InitModel)
            Me.__Model = value
            RaisePropertyChanged("Model")
        End Set
    End Property


    'この画面に／から遷移するフラグ
    'ウィンドウ側で購読する
    Private _InitFlag As Boolean
    Public Property InitFlag As Boolean
        Get
            Return Me._InitFlag
        End Get
        Set(value As Boolean)
            Me._InitFlag = value
            RaisePropertyChanged("InitFlag")
        End Set
    End Property


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

    Private _ConnectionString As String
    Public Property ConnectionString As String
        Get
            Return Me._ConnectionString
        End Get
        Set(value As String)
            Me._ConnectionString = value
            RaisePropertyChanged("ConnectionString")
            Call Me._CheckInitEnableFlag()
        End Set
    End Property


    Private __OtherProperty As String
    Private ReadOnly Property _OtherProperty As String
        Get
            Return _GetOtherProperty()
        End Get
    End Property


    Private Function _GetOtherProperty() As String
        Dim txt As String
        txt = InitViewModel.DEFUALT_OTHERPROPERTY
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
        _GetOtherProperty = txt
    End Function

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

    Private Sub _CheckInitEnableFlag()
        Dim b As Boolean : b = True
        If Me.ServerName = vbNullString Then
            b = False
        End If
        If Me.DataBaseName = vbNullString Then
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
        Dim sql As New SqlModel
        Dim pm As New ProjectModel
        Dim proxy As ProjectCheckProxy

        sql.ServerName = Me.ServerName
        sql.DataBaseName = Me.DataBaseName
        sql.ConnectionString = Me.ConnectionString

        Call sql.AccessTest()

        If sql.AccessResult Then
            Me.InitFlag = True
        End If

        Me._Model.ServerName = Me.ServerName
        Me._Model.DataBaseName = Me.DataBaseName
        Me._Model.ConnectionString = Me.ConnectionString
        Me._Model.InitFlag = Me.InitFlag

        proxy = AddressOf pm.FileCheck
        Select Case pm.ProjectCheck(proxy, INITVIEW_FILE)
            Case 0
                pm.ModelSave(Of InitModel)(INITVIEW_FILE, Me._Model)
            Case 1
                pm.FileEstablish(INITVIEW_FILE)
                pm.ModelSave(Of InitModel)(INITVIEW_FILE, Me._Model)
            Case 99
            Case 999
        End Select

    End Sub

    Private Function _InitCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._InitEnableFlag
    End Function

    Private Sub _SetDefault()
        Me.ServerName = InitViewModel.DEFAULT_SERVER_NAME
        Me.DataBaseName = InitViewModel.DEFAULT_DATABASE_NAME
    End Sub

    Sub New()
        Dim pm As New ProjectModel
        Dim proxy As ProjectCheckProxy
        Dim proxy2 As ModelLoadProxy(Of InitModel)

        proxy = AddressOf pm.FileCheck
        proxy2 = AddressOf pm.ModelLoad(Of InitModel)

        Select Case pm.ProjectCheck(proxy, INITVIEW_FILE)
            Case 0
                Me._Model = proxy2(INITVIEW_FILE)
                If Me._Model IsNot Nothing Then
                    Me.ServerName = Me._Model.ServerName
                    If String.IsNullOrEmpty(Me.ServerName) Then
                        Me.ServerName = InitViewModel.DEFAULT_SERVER_NAME
                    End If

                    Me.DataBaseName = Me._Model.DataBaseName
                    If String.IsNullOrEmpty(Me.DataBaseName) Then
                        Me.DataBaseName = InitViewModel.DEFAULT_DATABASE_NAME
                    End If

                    Me.ConnectionString = Me._Model.ConnectionString
                    If String.IsNullOrEmpty(Me.ConnectionString) Then
                        Me.ConnectionString = InitViewModel.DEFAULT_CONNECTIONSTRING
                    End If

                    Me.InitFlag = Me._Model.InitFlag
                    If Me._Model.InitFlag = Nothing Then
                        Me.InitFlag = False
                    End If
                Else
                    Me._Model = New InitModel
                    Call Me._SetDefault()
                End If
            Case Else
                Me._Model = New InitModel
                Call Me._SetDefault()
        End Select
    End Sub
End Class
