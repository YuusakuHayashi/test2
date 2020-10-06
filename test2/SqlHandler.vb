Imports System.Data
Imports System.Data.SqlClient

Public Class SqlHandler
    Public ServerVersion As String

    Private _ConnectionString As String
    Public Property ConnectionString As String
        Get
            Return Me._ConnectionString
        End Get
        Set(value As String)
            Me._ConnectionString = value
        End Set
    End Property

    ' 接続結果メッセージ
    Private _Query As String
    Public Property Query As String
        Get
            Return Me._Query
        End Get
        Set(value As String)
            Me._Query = value
        End Set
    End Property

    ' クエリ結果
    Private _Result As DataSet
    Public Property Result As DataSet
        Get
            Return Me._Result
        End Get
        Set(value As DataSet)
            Me._Result = value
        End Set
    End Property

    ' 接続結果
    Private _ResultFlag As Boolean
    Public Property ResultFlag As Boolean
        Get
            Return Me._ResultFlag
        End Get
        Set(value As Boolean)
            Me._ResultFlag = value
        End Set
    End Property

    ' 実行結果
    Private _QueryResult As Integer
    Public Property QueryResult As Integer
        Get
            Return Me._QueryResult
        End Get
        Set(value As Integer)
            Me._QueryResult = value
        End Set
    End Property

    ' 接続結果メッセージ
    Private _ResultMessage As String
    Public Property ResultMessage As String
        Get
            Return Me._ResultMessage
        End Get
        Set(value As String)
            Me._ResultMessage = value
        End Set
    End Property

    ' 実行結果エラー
    Private _ErrorResult As SqlException
    Public Property ErrorResult As SqlException
        Get
            Return Me._ErrorResult
        End Get
        Set(value As SqlException)
            Me._ErrorResult = value
        End Set
    End Property


    ' ＳＱＬＣｏｎｎｅｃｔｉｏｎ
    Private __Connection As SqlConnection
    Public Property _Connection As SqlConnection
        Get
            Return Me.__Connection
        End Get
        Set(value As SqlConnection)
            Me.__Connection = value
        End Set
    End Property

    ' ＳＱＬＴｒａｎｓａｃｔｉｏｎ
    Private __Transaction As SqlTransaction
    Public Property _Transaction As SqlTransaction
        Get
            Return Me.__Transaction
        End Get
        Set(value As SqlTransaction)
            Me.__Transaction = value
        End Set
    End Property

    ' ＳＱＬＣｏｍｍａｎｄ
    Private __Command As SqlCommand
    Public Property _Command As SqlCommand
        Get
            Return Me.__Command
        End Get
        Set(value As SqlCommand)
            Me.__Command = value
        End Set
    End Property

    '----------------------------------------------------------------------------------------------
    ' 例外をＳＱＬ例外か判定
    Private Sub _GetSqlError(ByVal ex As Exception)
        Dim b As Boolean
        b = TypeOf ex Is SqlException
        If b Then
            Me.ErrorResult = CType(ex, SqlException)
        Else
            Me.ErrorResult = Nothing
        End If
    End Sub
    '----------------------------------------------------------------------------------------------

    '----------------------------------------------------------------------------------------------
    Public Sub DeleteTable(ByVal tbl)
        Dim query As String
        query = $"  IF OBJECT_ID('{tbl}', 'U') IS NOT NULL"
        query &= $"    BEGIN DROP TABLE {tbl}"
        query &= $" END"
        Call Execute(query)
    End Sub
    '----------------------------------------------------------------------------------------------

    '----------------------------------------------------------------------------------------------
    ' 汎用のＳＱＬ実行

    ' 一連のＳＱＬ実行を行う（ＳＱＬ文はユーザ指定）
    Public Overloads Sub Execute(ByVal query As String)
        _Main(
            AddressOf _ConnectionStart,
            query,
            AddressOf _ExecuteMyQuery,
            AddressOf _ConnectionCommit,
            AddressOf _ConnectionFailed,
            AddressOf _ConnectionClose
        )
    End Sub

    ' 一連のＳＱＬ実行を行う（実行内容はユーザ指定）
    Public Overloads Sub Execute(ByRef proxy As ConnectionExecuteProxy)
        _Main(
            AddressOf _ConnectionStart,
            proxy,
            AddressOf _ConnectionCommit,
            AddressOf _ConnectionFailed,
            AddressOf _ConnectionClose
        )
    End Sub

    ' 汎用のアクセステスト
    'Public Overloads Sub AccessTest()
    '    _Main(
    '        AddressOf _ConnectionStart,
    '        AddressOf _NoDataGet,
    '        AddressOf _ConnectionCommit,
    '        AddressOf _ConnectionFailed,
    '        AddressOf _ConnectionClose
    '    )
    'End Sub

    ' 汎用のアクセステスト
    Public Overloads Function AccessTest() As Boolean
        Call Me._Main(AddressOf _ConnectionStart,
                      AddressOf _NoDataGet,
                      AddressOf _ConnectionCommit,
                      AddressOf _ConnectionFailed,
                      AddressOf _ConnectionClose)
        AccessTest = Me.ResultFlag
    End Function
    '----------------------------------------------------------------------------------------------


    ' ExecuteAccessProxy ------------------------------------------------------------------------------
    ' クエリ実行メソッド関連

    Public Delegate Sub ConnectionExecuteProxy()
    Public Delegate Sub ConnectionExecuteProxyWithQuery(ByVal query As String)

    ' コマンドプロパティをセットします
    Private Sub _SetCommand(ByVal txt As String, ByVal timeout As Integer)
        Try
            Me._Command = New System.Data.SqlClient.SqlCommand
            Me._Command = Me._Connection.CreateCommand()
            Me._Command.CommandText = txt
            Me._Command.CommandType = CommandType.Text
            Me._Command.CommandTimeout = timeout
            Me._Command.Transaction = Me._Transaction
        Catch ex As Exception
            Call _GetSqlError(ex)
            Me.ResultMessage = ex.Message
            Throw New Exception(ex.Message, ex)
        Finally
        End Try
    End Sub


    ' 基本的なクエリ実行メソッドです
    Private Overloads Sub _ExecuteMyQuery()
        Try
            Call _SetCommand(Me.Query, 30)
            Me.QueryResult = Me._Command.ExecuteNonQuery()
            'Me._Transaction.Commit()
        Catch ex As Exception
            Call _GetSqlError(ex)
            Me.ResultMessage = ex.Message
            Throw New Exception(ex.Message, ex)
        Finally
        End Try
    End Sub

    ' _ExecuteMyQuery() を外部公開したものです
    Public Overloads Sub ExecuteMyQuery()
        _ExecuteMyQuery()
    End Sub

    ' _ExecuteMyQuery() のクエリを外部から指定可能にしたものです
    Private Overloads Sub _ExecuteMyQuery(ByVal query As String)
        Try
            Call _SetCommand(query, 30)
            Me.QueryResult = Me._Command.ExecuteNonQuery()
        Catch ex As Exception
            Call _GetSqlError(ex)
            Me.ResultMessage = ex.Message
            Throw New Exception(ex.Message, ex)
        Finally
        End Try
    End Sub

    ' _ExecuteMyQuery(query) を外部公開したものです
    Public Overloads Sub ExecuteMyQuery(ByVal query As String)
        _ExecuteMyQuery(query)
    End Sub

    ' データセットを取得するクエリ実行メソッドです
    Private Overloads Sub _ExecuteMyGetQuery()
        Dim sda As System.Data.SqlClient.SqlDataAdapter
        Dim ds As DataSet
        Try
            Call _SetCommand(Me.Query, 30)

            sda = New System.Data.SqlClient.SqlDataAdapter(Me._Command)
            ds = New System.Data.DataSet
            sda.Fill(ds)

            Me.Result = ds
        Catch ex As Exception
            Call _GetSqlError(ex)
            Me.ResultMessage = ex.Message
            Throw New Exception(ex.Message)
        Finally
            If sda IsNot Nothing Then
                sda.Dispose()
            End If
            If ds IsNot Nothing Then
                ds.Dispose()
            End If
        End Try
    End Sub

    ' _ExecuteMyGetQuery() を外部公開したものです
    Public Overloads Sub ExecuteMyGetQuery()
        _ExecuteMyGetQuery()
    End Sub


    ' _ExecuteMyGetQuery() のクエリを外部から指定可能にしたものです
    Private Overloads Sub _ExecuteMyGetQuery(ByVal query As String)
        Dim sda As System.Data.SqlClient.SqlDataAdapter
        Dim ds As DataSet
        Try
            Call _SetCommand(Me.Query, 30)

            sda = New System.Data.SqlClient.SqlDataAdapter(Me._Command)
            ds = New System.Data.DataSet
            sda.Fill(ds)

            Me.Result = ds
        Catch ex As Exception
            Call _GetSqlError(ex)
            Me.ResultMessage = ex.Message
            Throw New Exception(ex.Message)
        Finally
            If sda IsNot Nothing Then
                sda.Dispose()
            End If
            If ds IsNot Nothing Then
                ds.Dispose()
            End If
        End Try
    End Sub

    ' _ExecuteMyGetQuery(query) を外部公開したものです
    Public Overloads Sub ExecuteMyGetQuery(ByVal query As String)
        _ExecuteMyGetQuery(query)
    End Sub


    ' クエリ実行しないクエリ実行メソッドです
    Private Overloads Sub _NoDataGet()
        Try
            Call _SetCommand(vbNullString, 30)
        Catch ex As Exception
            Call _GetSqlError(ex)
            Me.ResultMessage = ex.Message
            Throw New Exception(ex.Message, ex)
        Finally
        End Try
    End Sub
    '--------------------------------------------------------------------------------------------------

    '-- ConnectionCommitProxy -------------------------------------------------------------------------
    ' コミットメソッド関係
    Public Delegate Sub ConnectionCommitProxy()

    ' 実行結果をコミットするメソッドです
    Private Sub _ConnectionCommit()
        Try
            Me._Transaction.Commit()
        Catch ex As Exception
            Me.ResultMessage = ex.Message
            Throw New Exception(ex.Message)
        Finally
        End Try
    End Sub

    ' _ConnectionCommit() を外部公開したものです
    Public Sub ConnectionCommit()
        Call Me._ConnectionCommit()
    End Sub
    '--------------------------------------------------------------------------------------------------


    '-- ConnectionStartProxy --------------------------------------------------------------------------
    ' 接続開始メソッド関係
    Public Delegate Sub ConnectionStartProxy()

    ' 接続を開始するメソッドです
    Private Sub _ConnectionStart()
        Try
            Me._Connection = New SqlConnection(Me.ConnectionString)
            Me._Connection.Open()
            Me._Transaction = Me._Connection.BeginTransaction()
        Catch ex As Exception
            Me.ResultMessage = ex.Message
            Throw New Exception(ex.Message)
        Finally
        End Try
    End Sub

    Public Sub ConnectionStart()
        Call Me._ConnectionStart()
    End Sub
    '--------------------------------------------------------------------------------------------------


    '-- ConnectionFailedProxy -------------------------------------------------------------------------
    Public Delegate Sub ConnectionFailedProxy()

    Private Sub _ConnectionFailed()
        Try
            If Me._Transaction IsNot Nothing Then
                Me._Transaction.Rollback()
            End If
        Catch ex As Exception
            Me.ResultMessage = ex.Message
            Throw New Exception(ex.Message)
        Finally
        End Try
    End Sub

    Public Sub ConnectionFailed()
        _ConnectionFailed()
    End Sub
    '--------------------------------------------------------------------------------------------------


    '-- ConnectionCloseProxy --------------------------------------------------------------------------
    Public Delegate Sub ConnectionCloseProxy()

    Private Sub _ConnectionClose()
        If Me._Command IsNot Nothing Then
            Me._Command.Dispose()
        End If
        If Me._Transaction IsNot Nothing Then
            Me._Transaction.Dispose()
        End If
        If Me._Connection IsNot Nothing Then
            Me._Connection.Close()
            Me._Connection.Dispose()
        End If
    End Sub

    Public Sub ConnectionClose()
        _ConnectionClose()
    End Sub
    '--------------------------------------------------------------------------------------------------


    Private Overloads Sub _Main(ByRef startProxy As ConnectionStartProxy,
                                ByRef executeProxy As ConnectionExecuteProxy,
                                ByRef commitProxy As ConnectionCommitProxy,
                                ByRef failedProxy As ConnectionFailedProxy,
                                ByRef closeProxy As ConnectionCloseProxy)
        Dim b As Boolean : b = False
        Try
            ' アクセス開始
            startProxy()

            Me.ServerVersion = Me._Connection.ServerVersion

            executeProxy()
            commitProxy()

            ' Success
            b = True
        Catch ex As Exception
            ' アクセス失敗
            failedProxy()
        Finally
            Me._ResultFlag = b
            ' アクセス終了
            closeProxy()
        End Try
    End Sub


    Private Overloads Sub _Main(ByRef startProxy As ConnectionStartProxy,
                                ByVal query As String,
                                ByRef executeProxy As ConnectionExecuteProxyWithQuery,
                                ByRef commitProxy As ConnectionCommitProxy,
                                ByRef failedProxy As ConnectionFailedProxy,
                                ByRef closeProxy As ConnectionCloseProxy)
        Dim b As Boolean : b = False
        Try
            ' アクセス開始
            startProxy()

            Me.ServerVersion = Me._Connection.ServerVersion

            executeProxy(query)
            commitProxy()

            ' Success
            b = True
        Catch ex As Exception
            ' アクセス失敗
            failedProxy()
        Finally
            Me._ResultFlag = b
            ' アクセス終了
            closeProxy()
        End Try
    End Sub

End Class
