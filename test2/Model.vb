﻿Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.ObjectModel

Public Class Model
    Inherits ProjectBaseModel(Of Model)

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

    ' 履歴 ------------------------------------------------------------'
    Private _History As HistoryModel
    Public Property History As HistoryModel
        Get
            Return Me._History
        End Get
        Set(value As HistoryModel)
            Me._History = value
            RaisePropertyChanged("History")
        End Set
    End Property
    '------------------------------------------------------------------'


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


    ' 接続結果メッセージ
    Private _AccessMessage As String
    Public Property AccessMessage As String
        Get
            Return Me._AccessMessage
        End Get
        Set(value As String)
            Me._AccessMessage = value
        End Set
    End Property


    ' 接続のみ確認する
    Public Sub AccessTest()
        ' クエリ文は不要
        Me.Query = vbNullString

        ' 取得するＤＳは不要
        Call Me._DataBaseAccess(AddressOf Me._GetNoDataSet)
    End Sub




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

    Private Function _GetNoDataSet(ByVal scmd As SqlCommand) As DataSet
        _GetNoDataSet = New DataSet
    End Function


    Public Function GetSqlDataSet(ByVal scmd As SqlCommand) As DataSet
        GetSqlDataSet = _GetSqlDataSet(scmd)
    End Function

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
            Throw New Exception(ex.Message)
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
    Public Sub DataBaseAccess(ByRef proxy As GetDataSetProxy)
        Call _DataBaseAccess(proxy)
    End Sub


    ' データベースアクセス
    Private Sub _DataBaseAccess(ByRef proxy As GetDataSetProxy)
        Dim scmd As SqlCommand
        Dim scon As SqlConnection
        Dim strn As SqlTransaction
        Dim ds As DataSet

        Me._AccessResult = False
        Me.AccessMessage = vbNullString

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
            Finally
                Me.AccessMessage = ex.Message
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



    Public Overrides Sub MemberCheck()
        If History Is Nothing Then
            History = New HistoryModel
        End If

        If Server Is Nothing Then
            Server = New ServerModel With {
                .Name = "No Server",
                .IsEnabled = False,
                .IsChecked = False,
                .DataBases = New ObservableCollection(Of DataBaseModel) From {
                    New DataBaseModel With {
                        .Name = "No DataBase",
                        .IsEnabled = False,
                        .IsChecked = False,
                        .DataTables = New ObservableCollection(Of DataTableModel) From {
                            New DataTableModel With {
                                .Name = "No DataTable",
                                .IsEnabled = False,
                                .IsChecked = False
                            }
                        }
                    }
                }
            }
        End If
        If Server.DataBases Is Nothing Then
            Server.DataBases = New ObservableCollection(Of DataBaseModel) From {
                New DataBaseModel With {
                    .Name = "No DataBase",
                    .IsEnabled = False,
                    .IsChecked = False,
                    .DataTables = New ObservableCollection(Of DataTableModel) From {
                        New DataTableModel With {
                            .Name = "No DataTable",
                            .IsEnabled = False,
                            .IsChecked = False
                        }
                    }
                }
            }
        End If
        If Server.DataBases.Count = 0 Then
            Server.DataBases.Add(New DataBaseModel With {
                .Name = "No DataBase",
                .IsEnabled = False,
                .IsChecked = False,
                .DataTables = New ObservableCollection(Of DataTableModel) From {
                    New DataTableModel With {
                        .Name = "No DataTable",
                        .IsEnabled = False,
                        .IsChecked = False
                    }
                }
            })
        End If
        If Server.DataBases(0).DataTables.Count = 0 Then
            Server.DataBases(0).DataTables.Add(New DataTableModel With {
                .Name = "No DataTable",
                .IsChecked = False,
                .IsEnabled = False
            })
        End If
    End Sub

    Sub New()
    End Sub
End Class
