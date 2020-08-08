Imports System.Data
Imports System.Data.SqlClient

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


    Public Sub AccessTest()
        Me.Query = vbNullString
        Call Me._DataBaseAccess(Me)
    End Sub

    Public Sub AccessTestProxy(ByVal b As Boolean)
        Me.AccessFlag = b
    End Sub

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

    Private Sub _DataBaseAccess(ByVal sql As SqlModel)
        Dim scmd As SqlCommand
        Dim scon As SqlConnection
        Dim strn As SqlTransaction

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
End Class
