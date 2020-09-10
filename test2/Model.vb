Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.ObjectModel

Public Class Model
    Inherits ProjectModel(Of Model)

    ' このモデルを生成したＪＳＯＮファイル
    Private _SourceFile As String
    Public Property SourceFile As String
        Get
            Return Me._SourceFile
        End Get
        Set(value As String)
            Me._SourceFile = value
        End Set
    End Property

    Private _ChangePageStrings As String()
    Public Property ChangePageStrings As String()
        Get
            Return Me._ChangePageStrings
        End Get
        Set(value As String())
            If value IsNot Nothing Then
                If value.Length = 2 Then
                    Me._ChangePageStrings = value
                    RaisePropertyChanged("ChangePageStrings")
                End If
            End If
        End Set
    End Property

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

    Private _ConnectionString As String
    Public Property ConnectionString As String
        Get
            Return Me._ConnectionString
        End Get
        Set(value As String)
            Me._ConnectionString = value
        End Set
    End Property

    Public OtherProperty As String


    Private _Server As ServerModel
    Public Property Server As ServerModel
        Get
            Return Me._Server
        End Get
        Set(value As ServerModel)
            Me._Server = value
            RaisePropertyChanged("Server")
        End Set
    End Property


    Public ServerVersion As String


    ' 接続関連 --------------------------------------------------------'

    ' クエリ
    Public Query As String


    ' クエリ結果
    Private _QueryResult As DataSet
    Public ReadOnly Property QueryResult As DataSet
        Get
            Return _QueryResult
        End Get
    End Property


    ' 接続結果
    Private _AccessResult As Boolean
    Public ReadOnly Property AccessResult As Boolean
        Get
            Return Me._AccessResult
        End Get
    End Property


    ' 接続のみ確認する
    Public Sub AccessTest()
        ' クエリ文は不要
        Me.Query = vbNullString

        ' 取得するＤＳは不要
        Call Me._DataBaseAccess(
            Function()
                Return New DataSet
            End Function)
    End Sub


    'Public ReloadDataBase
    'End Sub
    ' テーブル一覧の取得
    'Public Sub GetUserTables()
    '    Dim proxy As GetDataSetProxy
    '    Dim dt As DataTable

    '    Me.Query = "SELECT * FROM SYS.OBJECTS WHERE TYPE = 'U'"
    '    proxy = AddressOf Me._GetSqlDataSet
    '    Call Me._DataBaseAccess(proxy)

    '    dt = Me._QueryResult.Tables(0)

    '    Me.Server.DataBases(0).DataTables = Nothing
    '    Me.Server.DataBases(0).DataTables = New ObservableCollection(Of DataTableModel)

    '    For Each r In dt.Rows
    '        Me.Server.DataBases(0).DataTables.Add(New DataTableModel With {.Name = r("name")})
    '    Next
    'End Sub


    ' サーバー全体の更新


    ' サーバー全体の更新
    Public Sub ReLoadServer()
        Dim sm As ServerModel
        Dim proxy As GetDataSetProxy
        Dim dt As DataTable
        Dim dbs As ObservableCollection(Of DataBaseModel)
        Dim dts As ObservableCollection(Of DataTableModel)

        Me.Query = "SELECT * FROM SYS.OBJECTS WHERE TYPE = 'U'"
        proxy = AddressOf Me._GetSqlDataSet
        Call Me._DataBaseAccess(proxy)

        dt = Me._QueryResult.Tables(0)

        dts = New ObservableCollection(Of DataTableModel)
        For Each r In dt.Rows
            dts.Add(New DataTableModel With {
                .Name = r("name")
            })
        Next

        dbs = New ObservableCollection(Of DataBaseModel)
        dbs.Add(New DataBaseModel With {
            .Name = Me.DataBaseName,
            .DataTables = dts
        })

        sm = New ServerModel With {
            .Name = Me.ServerName,
            .DataBases = dbs,
            .IsChecked = False
        }

        Me.Server = sm

        dt.Dispose()
        dts = Nothing
        dbs = Nothing
        sm = Nothing
    End Sub


    ' データセット取得
    'Public Function GetSqlDataSet() As DataSet
    '    Dim proxy As GetDataSetProxy

    '    proxy = AddressOf Me._GetSqlDataSet
    '    Call Me._DataBaseAccess(proxy)

    '    GetSqlDataSet = Me._QueryResult
    'End Sub

    Private Function _GetSqlDataSet(ByVal scmd As SqlCommand) As DataSet
        Dim sda As System.Data.SqlClient.SqlDataAdapter
        Dim ds As DataSet
        Try
            sda = New System.Data.SqlClient.SqlDataAdapter(scmd)
            ds = New System.Data.DataSet
            sda.Fill(ds)

            _GetSqlDataSet = ds
        Catch ex As Exception

            _GetSqlDataSet = Nothing
            Throw New Exception
        Finally
            If sda IsNot Nothing Then
                sda.Dispose()
            End If
            If ds IsNot Nothing Then
                ds.Dispose()
            End If
        End Try
    End Function





    ' データベースアクセス
    Private Sub _DataBaseAccess(ByRef proxy As GetDataSetProxy)
        Dim scmd As SqlCommand
        Dim scon As SqlConnection
        Dim strn As SqlTransaction
        Dim ds As DataSet

        Me._AccessResult = False

        Try
            '接続開始
            scon = New SqlConnection(Me.ConnectionString)
            scon.Open()
            strn = scon.BeginTransaction()

            'コマンド作成
            scmd = New System.Data.SqlClient.SqlCommand
            scmd = scon.CreateCommand()
            scmd.CommandText = Me.Query
            scmd.CommandType = CommandType.Text
            scmd.CommandTimeout = 30
            scmd.Transaction = strn

            '--- EXECUTE --------------------------------'
            ds = proxy(scmd)
            Me._QueryResult = ds
            '--------------------------------------------'

            Me.ServerVersion = scon.ServerVersion

            Me._AccessResult = True

            '成功
            'Me.AccessFlag = True
        Catch ex As Exception
            '失敗
            'Me.AccessFlag = False

            Try
                If strn IsNot Nothing Then
                    strn.Rollback()
                End If
            Catch ex2 As Exception
                '
            End Try
        Finally
            If scmd IsNot Nothing Then
                scmd.Dispose()
            End If
            If strn IsNot Nothing Then
                strn.Dispose()
            End If
            If scon IsNot Nothing Then
                scon.Close()
                scon.Dispose()
            End If
        End Try
    End Sub


    Public Sub DataBaseAccess(ByRef proxy As GetDataSetProxy)
        Dim scmd As SqlCommand
        Dim scon As SqlConnection
        Dim strn As SqlTransaction
        Dim ds As DataSet

        Me._AccessResult = False

        Try
            '接続開始
            scon = New SqlConnection(Me.ConnectionString)
            scon.Open()
            strn = scon.BeginTransaction()

            'コマンド作成
            scmd = New System.Data.SqlClient.SqlCommand
            scmd = scon.CreateCommand()
            scmd.CommandText = Me.Query
            scmd.CommandType = CommandType.Text
            scmd.CommandTimeout = 30
            scmd.Transaction = strn

            '--- EXECUTE --------------------------------'
            ds = proxy(scmd)
            Me._QueryResult = ds
            '--------------------------------------------'

            Me.ServerVersion = scon.ServerVersion

            Me._AccessResult = True

            '成功
            'Me.AccessFlag = True
        Catch ex As Exception
            '失敗
            'Me.AccessFlag = False

            Try
                If strn IsNot Nothing Then
                    strn.Rollback()
                End If
            Catch ex2 As Exception
                '
            End Try
        Finally
            If scmd IsNot Nothing Then
                scmd.Dispose()
            End If
            If strn IsNot Nothing Then
                strn.Dispose()
            End If
            If scon IsNot Nothing Then
                scon.Close()
                scon.Dispose()
            End If
        End Try
    End Sub

    Public Sub MemberCheck()
        ''
        'If String.IsNullOrEmpty(Me.SourceFile) Then
        '    Me.SourceFile = vbNullString
        'End If

        ''
        'If Me.ChangePageStrings Is Nothing Then
        '    Me.ChangePageStrings = {vbNullString, vbNullString}
        'End If

        ''
        'If String.IsNullOrEmpty(Me.ServerName) Then
        '    Me.ServerName = vbNullString
        'End If

        ''
        'If String.IsNullOrEmpty(Me.DataBaseName) Then
        '    Me.DataBaseName = vbNullString
        'End If

        ''
        'If String.IsNullOrEmpty(Me.ConnectionString) Then
        '    Me.ConnectionString = vbNullString
        'End If

        ''
        'If String.IsNullOrEmpty(Me.OtherProperty) Then
        '    Me.OtherProperty = vbNullString
        'End If

        ''
        'If String.IsNullOrEmpty(Me.ServerVersion) Then
        '    Me.ServerVersion = vbNullString
        'End If

        ''
        'If Me.Server Is Nothing Then
        '    Me.Server = New ServerModel
        'End If

        ''
        'If String.IsNullOrEmpty(Me.Query) Then
        '    Me.Query = vbNullString
        'End If

        'If Me._QueryResult Is Nothing Then
        '    Me._QueryResult = New DataSet
        'End If
    End Sub

    Sub New()
        Call Me.MemberCheck()
    End Sub
End Class
