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
    Private _Result As Boolean
    Public Property Result As Boolean
        Get
            Return Me._Result
        End Get
        Set(value As Boolean)
            Me._Result = value
        End Set
    End Property

    ' 接続結果メッセージ
    Private _Message As String
    Public Property Message As String
        Get
            Return Me._Message
        End Get
        Set(value As String)
            Me._Message = value
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
    Public Function AccessTest() As Boolean
        Me.Query = "SELECT GETDATE()"
        Call Execute()
        AccessTest = Me.Result
    End Function

    Public Sub DeleteTable(ByVal tbl)
        Me.Query = $"  IF OBJECT_ID('{tbl}', 'U') IS NOT NULL"
        Me.Query &= $"    BEGIN DROP TABLE {tbl}"
        Me.Query &= $" END"
        Call Execute()
    End Sub
    '----------------------------------------------------------------------------------------------

    ' ExecuteAccessProxy ------------------------------------------------------------------------------
    Private Delegate Sub ExecuteProxy()
    Public ExecuteHandler As Action
    Public GetExecuteHandler As Func(Of DataSet, Object)
    Public ExecuteIfFailedHandler As Action
    Public ExecuteAfterRollbackIfFailedHandler As Action

    Private Overloads Sub _ExecuteQuery()
        Try
            Call Me._Command.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Overloads Function _GetDataSet() As DataSet
        Dim sda As System.Data.SqlClient.SqlDataAdapter
        Dim ds As DataSet : ds = Nothing
        Try
            sda = New System.Data.SqlClient.SqlDataAdapter(Me._Command)
            ds = New System.Data.DataSet
            sda.Fill(ds)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            If sda IsNot Nothing Then
                sda.Dispose()
            End If
            If ds IsNot Nothing Then
                ds.Dispose()
            End If
            _GetDataSet = ds
        End Try
    End Function

    ' クエリ実行しないクエリ実行メソッドです
    Private Overloads Sub _NoDataGet()
        Try
        Catch ex As Exception
            Call _GetSqlError(ex)
            Me.Message = ex.Message
            Throw New Exception(ex.Message, ex)
        Finally
        End Try
    End Sub

    Private Sub _BeforeExecute()
        Try
            Me._Connection = New SqlConnection(Me.ConnectionString)
            Me._Connection.Open()
            Me._Transaction = Me._Connection.BeginTransaction()

            Me._Command = New System.Data.SqlClient.SqlCommand
            Me._Command = Me._Connection.CreateCommand()
            Me._Command.CommandText = Me.Query
            Me._Command.CommandType = CommandType.Text
            Me._Command.CommandTimeout = 30
            Me._Command.Transaction = Me._Transaction
        Catch ex As Exception
            Me.Message = ex.Message
            Throw New Exception(ex.Message)
        Finally
        End Try
    End Sub

    Private Sub _AfterExecute()
        Try
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
        Catch ex As Exception
            Me.Message = ex.Message
            Throw New Exception(ex.Message)
        Finally
        End Try
    End Sub

    Public Sub Execute()
        Dim b = False
        Dim msg = vbNullString
        Dim a = Me.ExecuteIfFailedHandler
        Dim a2 = Me.ExecuteAfterRollbackIfFailedHandler

        Try
            Call _BeforeExecute()
            Call Me._Command.ExecuteNonQuery()
            Me._Transaction.Commit()
            b = True
        Catch ex As Exception
            msg = ex.Message
            Try
                If a <> Nothing Then
                    Call a()
                End If
                Me._Transaction.Rollback()
            Catch ex2 As Exception
                msg = ex2.Message
            End Try
            If a2 <> Nothing Then
                Call a2()
            End If
        Finally
            Call _AfterExecute()
            Me.Result = b
        End Try
    End Sub

    Public Overloads Function GetExecute() As Object
        Dim obj As Object : obj = Nothing
        Dim ds As DataSet : ds = Nothing
        Dim b = False
        Dim msg = vbNullString
        Dim f = Me.GetExecuteHandler
        Dim e2 = Me.ExecuteIfFailedHandler
        Dim e3 = Me.ExecuteAfterRollbackIfFailedHandler
        Dim sda As System.Data.SqlClient.SqlDataAdapter

        Try
            Call _BeforeExecute()

            sda = New System.Data.SqlClient.SqlDataAdapter(Me._Command)
            ds = New System.Data.DataSet
            sda.Fill(ds)

            obj = IIf(f <> Nothing, f(ds), ds)
            b = True
        Catch ex As Exception
            msg = ex.Message
            Try
                If e2 <> Nothing Then
                    Call e2()
                End If
                Me._Transaction.Rollback()
            Catch ex2 As Exception
                msg = ex2.Message
            End Try
            If e3 <> Nothing Then
                Call e3()
            End If
        Finally
            If sda IsNot Nothing Then
                sda.Dispose()
            End If
            If ds IsNot Nothing Then
                ds.Dispose()
            End If

            Call _AfterExecute()
            Me.Message = msg
            GetExecute = obj
            Me.Result = b
        End Try
    End Function
End Class
