Imports System.Data
Imports System.Data.SqlClient


Public Class SqlHandlerModel : Inherits BaseModel(Of SqlHandlerModel)


    Private _SqlM As SqlModel
    Public Property SqlM As SqlModel
        Get
            Return _SqlM
        End Get
        Set(value As SqlModel)
            _SqlM = value
        End Set
    End Property


    'Private _TreeViewM As TreeViewModel
    Public WriteOnly Property TreeViewM As TreeViewModel
        'Get
        '    Return _TreeViewM
        'End Get
        Set(value As TreeViewModel)
            '_TreeViewM = value
        End Set
    End Property



    'Public Sub AccessTest()
    '    Me.Query = vbNullString
    '    Call Me._DataBaseAccess(Me.SqlM)
    'End Sub

    Public Query As String
    Public ServerVersion As String


    'Private _AccessFlag As Boolean
    'Public Property AccessFlag As Boolean
    '    Get
    '        Return _AccessFlag
    '    End Get
    '    Set(value As Boolean)
    '        _AccessFlag = value
    '        RaisePropertyChanged("AccessFlag")
    '    End Set
    'End Property

    'Private _ConnectionText As String
    'Public ReadOnly Property ConnectionText(ByVal sql As SqlModel) As String
    '    Get
    '        Return "Data Source=" & sql.ServerName & ";" _
    '            & "Initial Catalog=" & sql.DataBaseName & ";" _
    '            & "Integrated Security=True;" _
    '            & "Connect Timeout=05;"
    '    End Get
    'End Property


#Region "データテーブル一覧"
    Private __DS As DataSet
    Private Property _DS As DataSet
        Get
            Return __DS
        End Get
        Set(value As DataSet)
            __DS = value
        End Set
    End Property
#End Region

#Region "データベースへのアクセス"

    'SQLコマンドをセットする必要があるが、空のデータセットを返す
    Private Function _NoDataSets(scmd) As DataSet
        _NoDataSets = New DataSet
    End Function



    'SQLコマンドからデータセットを取得する
    Private Function _GetDataSets(scmd) As DataSet
        Dim sda As SqlDataAdapter
        Dim ds As New DataSet
        Try
            sda = New SqlDataAdapter(scmd)
            sda.Fill(ds)

            _GetDataSets = ds
        Catch ex As Exception
            Throw ex
        Finally
            If sda IsNot Nothing Then
                sda.Dispose()
            End If
            If ds IsNot Nothing Then
                ds.Dispose()
            End If
        End Try
    End Function


    '接続のみ
    'Public Sub AccessTest()
    '    Me.Query = vbNullString
    '    Call Me._DataBaseAccess(Me.SqlM, AddressOf Me._NoDataSets)
    'End Sub


    'テーブル一覧の取得
    'Public Sub GetTables()
    '    Dim ltvm As New List(Of TreeViewModel)
    '    Dim parent As New TreeViewModel
    '    Dim child As New TreeViewModel
    '    Dim dt As DataTable


    '    Me.Query = "SELECT * FROM SYS.OBJECTS WHERE TYPE = 'U'"
    '    Call Me._DataBaseAccess(Me.SqlM, AddressOf Me._GetDataSets)


    '    dt = Me._DS.Tables(0)


    '    'TreeViewを構成
    '    parent = New TreeViewModel With {.RealName = Me.SqlM.ServerName}
    '    child = New TreeViewModel With {.RealName = Me.SqlM.DataBaseName}
    '    For i As Integer = 0 To dt.Rows.Count - 1
    '        child.Child.Add(New TreeViewModel With {.RealName = dt.Rows(i)("name")})
    '    Next
    '    parent.Child.Add(child)


    '    'TreeViewのリセット
    '    Me.TreeViewM = parent
    'End Sub


    'メインメソッド
    'Private Sub _DataBaseAccess(ByRef sql As SqlModel,
    '                            ByRef proxy As GetDataSetProxy)

    '    'proxy ... Func(SqlCommand) => DataSet

    '    Dim scmd As SqlCommand
    '    Dim scon As SqlConnection
    '    Dim strn As SqlTransaction
    '    Dim ds As DataSet

    '    Try
    '        '接続開始
    '        scon = New SqlConnection(Me.SqlM.ConnectionText)
    '        scon.Open()
    '        strn = scon.BeginTransaction()

    '        'コマンド作成
    '        scmd = New System.Data.SqlClient.SqlCommand
    '        scmd = scon.CreateCommand()
    '        scmd.CommandText = Me.Query
    '        scmd.CommandType = CommandType.Text
    '        scmd.CommandTimeout = 30
    '        scmd.Transaction = strn


    '        '--- EXECUTE --------------------------------'
    '        ds = proxy(scmd)
    '        Me._DS = ds
    '        '--------------------------------------------'


    '        sql.ServerVersion = scon.ServerVersion


    '        '成功
    '        'Me.AccessFlag = True
    '    Catch ex As Exception
    '        '失敗
    '        'Me.AccessFlag = False

    '        Try
    '            If strn IsNot Nothing Then
    '                strn.Rollback()
    '            End If
    '        Catch ex2 As Exception
    '            '
    '        End Try
    '    Finally
    '        If scmd IsNot Nothing Then
    '            scmd.Dispose()
    '        End If
    '        If strn IsNot Nothing Then
    '            strn.Dispose()
    '        End If
    '        If scon IsNot Nothing Then
    '            scon.Close()
    '            scon.Dispose()
    '        End If
    '    End Try
    'End Sub




#End Region

    Sub New(ByRef sm As SqlModel, ByRef tvm As TreeViewModel)
        Me.SqlM = sm
        Me.TreeViewM = tvm
    End Sub

    'Public Sub AccessTestProxy(ByVal b As Boolean)
    '    Me.AccessFlag = b
    'End Sub

End Class
