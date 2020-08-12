Imports System.Data
Imports System.Data.SqlClient

Delegate Function GetDataSetProxy(ByVal scmd As SqlCommand) As DataSet

Public Class SqlModel : Inherits BaseModel

    Private Property _ServerName As String
    Public Property ServerName As String
        Get
            Return _ServerName
        End Get
        Set(value As String)
            _ServerName = value
            RaisePropertyChanged("ServerName")
        End Set
    End Property


    Private _DataBaseName As String
    Public Property DataBaseName As String
        Get
            Return _DataBaseName
        End Get
        Set(value As String)
            _DataBaseName = value
            RaisePropertyChanged("DataBaseName")
        End Set
    End Property


    Private _DataTableName As String
    Public Property DataTableName As String
        Get
            Return _DataTableName
        End Get
        Set(value As String)
            _DataTableName = value
            RaisePropertyChanged("DataTableName")
        End Set
    End Property


    Private _FieldName As String
    Public Property FieldName As String
        Get
            Return _FieldName
        End Get
        Set(value As String)
            _FieldName = value
            RaisePropertyChanged("DataTableName")
        End Set
    End Property

    Private _SourceValue As String
    Public Property SourceValue As String
        Get
            Return _SourceValue
        End Get
        Set(value As String)
            _SourceValue = value
            RaisePropertyChanged("SourceValue")
        End Set
    End Property

    Private _DistinationValue As String
    Public Property DistinationValue As String
        Get
            Return _DistinationValue
        End Get
        Set(value As String)
            _DistinationValue = value
            RaisePropertyChanged("DistinationValue")
        End Set
    End Property

    Public Query As String
    Public ServerVersion As String


    Private _AccessFlag As Boolean
    Public Property AccessFlag As Boolean
        Get
            Return _AccessFlag
        End Get
        Set(value As Boolean)
            _AccessFlag = value
            RaisePropertyChanged("AccessFlag")
        End Set
    End Property


    Private _ConnectionText As String
    Public ReadOnly Property ConnectionText As String
        Get
            Return "Data Source=" & Me.ServerName & ";" _
                & "Initial Catalog=" & Me.DataBaseName & ";" _
                & "Integrated Security=True;" _
                & "Connect Timeout=05;"
        End Get
    End Property


    'Public Sub Save()
    '    Dim pm As New ProjectModel
    '    Dim cfm As New ConfigFileModel
    '    Dim f As String

    '    f = pm.LoadFileManagerFile.ConfigFileName
    '    cfm = pm.ModelLoad(Of ConfigFileModel)(f)
    '    cfm.SqlStatus = Me

    '    pm.ModelSave(Of ConfigFileModel)(f, cfm)

    '    pm = Nothing
    '    cfm = Nothing
    'End Sub

    Private _TreeViewM As TreeViewModel
    Public Property TreeViewM As TreeViewModel
        Get
            Return _TreeViewM
        End Get
        Set(value As TreeViewModel)
            _TreeViewM = value
        End Set
    End Property

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


    'Private __GetDataSetMethod As Func(Of SqlModel, DataSet)
    'Private Property _GetDataSetMethod As Func(Of SqlModel, DataSet)
    '    Get
    '        Return __GetDataSetMethod
    '    End Get
    '    Set(value As Func(Of SqlModel, DataSet))
    '        __GetDataSetMethod = value
    '    End Set
    'End Property



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
    Public Sub AccessTest()
        Me.Query = vbNullString
        Call Me._DataBaseAccess(Me, AddressOf Me._NoDataSets)
    End Sub


    'テーブル一覧の取得
    Public Sub GetTables()
        Dim ltvm As New List(Of TreeViewModel)
        Dim parent As New TreeViewModel
        Dim child As New TreeViewModel
        Dim dt As DataTable


        Me.Query = "SELECT * FROM SYS.OBJECTS WHERE TYPE = 'U'"
        Call Me._DataBaseAccess(Me, AddressOf Me._GetDataSets)


        dt = Me._DS.Tables(0)


        'TreeViewを構成
        parent = New TreeViewModel With {.RealName = Me.ServerName}
        child = New TreeViewModel With {.RealName = Me.DataBaseName}
        For i As Integer = 0 To dt.Rows.Count - 1
            child.Child.Add(New TreeViewModel With {.RealName = dt.Rows(i)("name")})
        Next
        parent.Child.Add(child)


        'TreeViewのリセット
        Me.TreeViewM = parent
    End Sub


    'メインメソッド
    Private Sub _DataBaseAccess(ByVal sql As SqlModel,
                                ByRef proxy As GetDataSetProxy)

        'proxy ... Func(SqlCommand) => DataSet

        Dim scmd As SqlCommand
        Dim scon As SqlConnection
        Dim strn As SqlTransaction
        Dim ds As DataSet

        Try
            '接続開始
            scon = New SqlConnection(Me.ConnectionText)
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
            Me._DS = ds
            '--------------------------------------------'


            sql.ServerVersion = scon.ServerVersion


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




#End Region


    Public Sub AccessTestProxy(ByVal b As Boolean)
        Me.AccessFlag = b
    End Sub

End Class
