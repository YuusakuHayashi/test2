Imports System.Data
Imports System.Data.SqlClient

Public Class Model : Inherits BaseModel
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
End Class
