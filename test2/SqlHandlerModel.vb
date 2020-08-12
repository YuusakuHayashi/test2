Imports System.Data
Imports System.Data.SqlClient

Public Class SqlHandlerModel : Inherits BaseModel

    Private _SqlM As SqlModel
    Public Property SqlM As SqlModel
        Get
            Return _SqlM
        End Get
        Set(value As SqlModel)
            _SqlM = value
        End Set
    End Property


    Private _TreeViewVM As TreeViewViewModel
    Public Property TreeViewVM As TreeViewViewModel
        Get
            Return _TreeViewVM
        End Get
        Set(value As TreeViewViewModel)
            _TreeViewVM = value
        End Set
    End Property



    Public Sub AccessTest()
        Me.Query = vbNullString
        Call Me._DataBaseAccess(Me.SqlM)
    End Sub

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


    Private _ConnectionText As String
    Public ReadOnly Property ConnectionText(ByVal sql As SqlModel) As String
        Get
            Return "Data Source=" & sql.ServerName & ";" _
                & "Initial Catalog=" & sql.DataBaseName & ";" _
                & "Integrated Security=True;" _
                & "Connect Timeout=05;"
        End Get
    End Property


    Private Sub _DataBaseAccess(ByVal sql As SqlModel)
        Dim scmd As SqlCommand
        Dim scon As SqlConnection
        Dim strn As SqlTransaction

        Try
            '接続開始
            scon = New SqlConnection(Me.ConnectionText(sql))
            scon.Open()
            strn = scon.BeginTransaction()

            'コマンド作成
            scmd = New System.Data.SqlClient.SqlCommand
            scmd = scon.CreateCommand()
            scmd.CommandText = Me.Query
            scmd.CommandType = CommandType.Text
            scmd.CommandTimeout = 30
            scmd.Transaction = strn


            Me.ServerVersion = scon.ServerVersion

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
