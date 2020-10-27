Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Model
    Inherits JsonHandler(Of Model)

    Private _Data As Object
    Public Property Data As Object
        Get
            Return Me._Data
        End Get
        Set(value As Object)
            Me._Data = value
        End Set
    End Property

    Public History As HistoryModel

    'Public Sub DataSave(ByVal project As ProjectInfoModel)
    '    Call Me.ModelSave(project.ModelFileName, Me)
    'End Sub

    '' AddProject時、Projectに必要なモデルをインスタンス化する
    'Public Sub Setup(ByVal project As ProjectInfoModel)
    '    '-- you henkou --------------------------------'
    '    Select Case project.Kind
    '        Case AppDirectoryModel.DBTEST
    '            Me.Data = New DBTestModel
    '        Case Else
    '    End Select
    '    '----------------------------------------------'
    'End Sub

    'Private Sub _DataInitialize(Of T As {New})()
    '    Dim obj As Object
    '    If Me.Data IsNot Nothing Then
    '        obj = Me.Data
    '        Select Case obj.GetType
    '            Case (New Object).GetType
    '                Me.Data = CType(obj, T)
    '            Case (New JObject).GetType
    '                ' Ｊｓｏｎからロードした場合は、JObject型になっている
    '                Me.Data = obj.ToObject(Of T)
    '            Case (New T).GetType
    '                Me.Data = obj
    '        End Select
    '    Else
    '        Me.Data = New T
    '    End If
    'End Sub

    'Public Sub Initialize(ByVal project As ProjectInfoModel)
    '    '-- you henkou --------------------------------'
    '    Select Case project.Kind
    '        Case AppDirectoryModel.DBTEST
    '            Call Me._DataInitialize(Of DBTestModel)()
    '        Case Else
    '    End Select
    '    '----------------------------------------------'
    'End Sub

    ' DBTestModel のオブジェクトの初期化
    'Private Sub _DBTestModelInitialize()
    '    If Me.Data Is Nothing Then
    '        Me.Data = New DBTestModel
    '    End If
    '    If Me.Data.Server Is Nothing Then
    '        Me.Data.Server = New ServerModel
    '    End If
    '    If Me.Data.History Is Nothing Then
    '        Me.Data.History = New HistoryModel
    '        Call Me.Data.History.NewLine("Historyは初期化されました")
    '    End If
    'End Sub

    '' このモデルを生成したＪＳＯＮファイル
    'Private _SourceFile As String
    'Public Property SourceFile As String
    '    Get
    '        Return Me._SourceFile
    '    End Get
    '    Set(value As String)
    '        Me._SourceFile = value
    '    End Set
    'End Property

    'Private _ServerName As String
    'Public Property ServerName As String
    '    Get
    '        Return Me._ServerName
    '    End Get
    '    Set(value As String)
    '        Me._ServerName = value
    '    End Set
    'End Property

    'Private _DataBaseName As String
    'Public Property DataBaseName As String
    '    Get
    '        Return Me._DataBaseName
    '    End Get
    '    Set(value As String)
    '        Me._DataBaseName = value
    '    End Set
    'End Property

    'Private _TestTableName As String
    'Public Property TestTableName As String
    '    Get
    '        Return Me._TestTableName
    '    End Get
    '    Set(value As String)
    '        Me._TestTableName = value
    '    End Set
    'End Property

    'Private _ConnectionString As String
    'Public Property ConnectionString As String
    '    Get
    '        Return Me._ConnectionString
    '    End Get
    '    Set(value As String)
    '        Me._ConnectionString = value
    '    End Set
    'End Property

    'Public OtherProperty As String


    'Private _Server As ServerModel
    'Public Property Server As ServerModel
    '    Get
    '        Return Me._Server
    '    End Get
    '    Set(value As ServerModel)
    '        Me._Server = value
    '        RaisePropertyChanged("Server")
    '    End Set
    'End Property

    'Private _MenuFolder As MenuFolderModel
    'Public Property MenuFolder As MenuFolderModel
    '    Get
    '        Return Me._MenuFolder
    '    End Get
    '    Set(value As MenuFolderModel)
    '        Me._MenuFolder = value
    '        RaisePropertyChanged("MenuFolder")
    '    End Set
    'End Property

    '' 履歴 ------------------------------------------------------------'
    'Private _History As HistoryModel
    'Public Property History As HistoryModel
    '    Get
    '        Return Me._History
    '    End Get
    '    Set(value As HistoryModel)
    '        Me._History = value
    '        RaisePropertyChanged("History")
    '    End Set
    'End Property
    ''------------------------------------------------------------------'


    ''Public ServerVersion As String


    '' 接続関連 --------------------------------------------------------'
    ''Public Query As String

    '' クエリ
    ''Public Query As String


    ''' クエリ結果
    ''Private _QueryResult As DataSet
    ''Public ReadOnly Property QueryResult As DataSet
    ''    Get
    ''        Return _QueryResult
    ''    End Get
    ''End Property


    ''' 接続結果
    ''Private _AccessResult As Boolean
    ''Public ReadOnly Property AccessResult As Boolean
    ''    Get
    ''        Return Me._AccessResult
    ''    End Get
    ''End Property


    ''' 接続結果メッセージ
    ''Private _AccessMessage As String
    ''Public Property AccessMessage As String
    ''    Get
    ''        Return Me._AccessMessage
    ''    End Get
    ''    Set(value As String)
    ''        Me._AccessMessage = value
    ''    End Set
    ''End Property


    '' 接続のみ確認する
    ''Public Sub AccessTest()
    ''    ' クエリ文は不要
    ''    Me.Query = vbNullString

    ''    ' 取得するＤＳは不要
    ''    Call Me._DataBaseAccess(AddressOf Me._GetNoDataSet)
    ''End Sub


    '' ＳＱＬ関係の汎用ロジック ------------------------------------------------------------------------
    'Public Function AccessTest() As MySql
    '    Dim mysql As New MySql

    '    mysql.ConnectionString = Me.ConnectionString

    '    mysql.AccessTest()

    '    AccessTest = mysql
    'End Function

    'Public Function GetUserTables() As MySql
    '    Dim mysql As New MySql

    '    mysql.ConnectionString = Me.ConnectionString
    '    mysql.Query = "SELECT * FROM SYS.OBJECTS WHERE TYPE = 'U'"
    '    mysql.Execute(AddressOf mysql.GetSqlData)

    '    GetUserTables = mysql
    'End Function

    '' サーバー全体の更新
    'Public Sub ReLoadServer()
    '    Dim sm As ServerModel
    '    Dim dt As DataTable
    '    Dim dbs As ObservableCollection(Of DataBaseModel)
    '    Dim dts As ObservableCollection(Of DataTableModel)
    '    Dim mysql As MySql

    '    mysql = GetUserTables()

    '    dt = mysql.Result.Tables(0)

    '    dts = New ObservableCollection(Of DataTableModel)
    '    For Each r In dt.Rows
    '        dts.Add(New DataTableModel With {
    '            .Name = r("name"),
    '            .IsEnabled = True,
    '            .IsChecked = True
    '        })
    '    Next

    '    dbs = New ObservableCollection(Of DataBaseModel)
    '    dbs.Add(New DataBaseModel With {
    '        .Name = Me.DataBaseName,
    '        .DataTables = dts,
    '        .IsEnabled = True,
    '        .IsChecked = True
    '    })

    '    sm = New ServerModel With {
    '        .Name = Me.ServerName,
    '        .DataBases = dbs,
    '        .IsEnabled = True,
    '        .IsChecked = True
    '    }

    '    Me.Server = sm

    '    dt.Dispose()
    '    dts = Nothing
    '    dbs = Nothing
    '    sm = Nothing
    'End Sub



    'Public Overloads Function ModelLoad(ByVal f As String) As Model
    '    Dim proxy As LoadedMethodProxy

    '    proxy = Sub(ByRef m As Model)
    '                If m.History Is Nothing Then
    '                    m.History = New HistoryModel
    '                End If
    '                m.History.AddLine(f & "が読み込まれました。")
    '            End Sub

    '    ModelLoad = ModelLoad(f, proxy)
    'End Function


    '' ウィンドウイニシャライズ時にメンバのチェックを行う
    '' セット後、メンバチェックを行いたいオブジェクトは、継承したMemberCheckメソッドを
    '' ここで呼び出す
    'Public Overrides Sub MemberCheck()
    '    If History Is Nothing Then
    '        History = New HistoryModel
    '    End If
    '    History.CheckLineCounts()

    '    If MenuFolder Is Nothing Then
    '        MenuFolder = New MenuFolderModel
    '    End If
    '    MenuFolder.MemberCheck()

    '    If Server Is Nothing Then
    '        Server = New ServerModel With {
    '            .Name = "No Server",
    '            .IsEnabled = True,
    '            .IsChecked = True,
    '            .DataBases = New ObservableCollection(Of DataBaseModel) From {
    '                New DataBaseModel With {
    '                    .Name = "No DataBase",
    '                    .IsEnabled = False,
    '                    .IsChecked = False,
    '                    .DataTables = New ObservableCollection(Of DataTableModel) From {
    '                        New DataTableModel With {
    '                            .Name = "No DataTable",
    '                            .IsEnabled = False,
    '                            .IsChecked = False
    '                        }
    '                    }
    '                }
    '            }
    '        }
    '    End If
    '    If Server.DataBases Is Nothing Then
    '        Server.DataBases = New ObservableCollection(Of DataBaseModel) From {
    '            New DataBaseModel With {
    '                .Name = "No DataBase",
    '                .IsEnabled = False,
    '                .IsChecked = False,
    '                .DataTables = New ObservableCollection(Of DataTableModel) From {
    '                    New DataTableModel With {
    '                        .Name = "No DataTable",
    '                        .IsEnabled = False,
    '                        .IsChecked = False
    '                    }
    '                }
    '            }
    '        }
    '    End If
    '    If Server.DataBases.Count = 0 Then
    '        Server.DataBases.Add(New DataBaseModel With {
    '            .Name = "No DataBase",
    '            .IsEnabled = False,
    '            .IsChecked = False,
    '            .DataTables = New ObservableCollection(Of DataTableModel) From {
    '                New DataTableModel With {
    '                    .Name = "No DataTable",
    '                    .IsEnabled = False,
    '                    .IsChecked = False
    '                }
    '            }
    '        })
    '    End If
    '    If Server.DataBases(0).DataTables.Count = 0 Then
    '        Server.DataBases(0).DataTables.Add(New DataTableModel With {
    '            .Name = "No DataTable",
    '            .IsChecked = False,
    '            .IsEnabled = False
    '        })
    '    End If
    'End Sub

    'Sub New()
    'End Sub
End Class
