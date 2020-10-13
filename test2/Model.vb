Imports System.ComponentModel

Public Class Model
    Inherits BaseModel(Of Model)

    Private _Data As Object
    Public Property Data As Object
        Get
            Return Me._Data
        End Get
        Set(value As Object)
            Me._Data = value
        End Set
    End Property

    '-- ディクショナリでのモデル保持は現在廃止 ---------------------------------------------------'
    '-- 使い道はありそうなので、残しておく -------------------------------------------------------'
    'Private _DataDictionary As Dictionary(Of String, Object)
    'Public Property DataDictionary As Dictionary(Of String, Object)
    '    Get
    '        Return Me._DataDictionary
    '    End Get
    '    Set(value As Dictionary(Of String, Object))
    '        Me._DataDictionary = value
    '    End Set
    'End Property

    '' そのモデルに対応するデータがない場合、
    '' その格納用のメモリを確保します
    'Private Sub _DictionaryCheck(ByVal view As String, ByVal nm As String)
    '    If Me.DataDictionary Is Nothing Then
    '        Me.DataDictionary = New Dictionary(Of String, Object)
    '    End If
    'End Sub

    '' モデルのディクショナリへの追加・更新を行います
    'Public Overloads Sub SetData(ByVal modelName As String, ByRef data As Object)
    '    If Me.DataDictionary Is Nothing Then
    '        Me.DataDictionary = New Dictionary(Of String, Object)
    '    End If
    '    If Not Me.DataDictionary.ContainsKey(modelName) Then
    '        Me.DataDictionary.Add(modelName, data)
    '    Else
    '        Me.DataDictionary(modelName) = data
    '    End If
    'End Sub
    '---------------------------------------------------------------------------------------------'

    Public Delegate Sub InitializeProxy(ByVal pk As String)

    Public Overloads Sub InitializeModels(ByVal project As ProjectInfoModel)
        Dim jh As New JsonHandler(Of Object)


        Dim cm As ConnectionModel
        Dim dbtm As DBTestModel
        Dim sm As ServerModel
        Dim hm As HistoryModel

        ModelFileName = project.ModelFileName

        '-- you henkou --------------------------------'
        Select Case project.Kind
            Case AppDirectoryModel.DB_TEST
                With jh
                    .LoadHandler = AddressOf ModelLoad(Of DBTestModel)
                    .LoadHandlerIfFailed = AddressOf NewLoad(Of DBTestModel)
                    .LoadHandlerIfNull = AddressOf NewLoad(Of DBTestModel)
                End With

                dbtm = jh.ModelLoad(Of DBTestModel)()
                dbtm = New DBTestModel
                Me.Data = dbtm
            Case Else
        End Select
        '----------------------------------------------'
    End Sub

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
