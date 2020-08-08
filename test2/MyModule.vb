''Imports Newtonsoft.Json
''Imports System.Data
''Imports System.ComponentModel
''Imports System.Threading
Imports System.Data.SqlClient

Module MyModule

    'Public Function TreeViewVMRecursive(ByRef ltvm As List(Of TreeViewModel)) As TreeViewViewModel
    '    Dim tvvm As New TreeViewViewModel
    '    Dim tvvm2 As TreeViewViewModel
    '    Dim l As List(Of TreeViewModel)

    '    If ltvm IsNot Nothing Then
    '        For Each tvm In ltvm
    '            tvvm.RealName = tvm.RealName
    '            tvvm.Checked = tvm.Checked
    '            '-- 直接再帰関数に.Childを入れるとウィルスチェックに掛かる ---
    '            l = tvm.Child

    '            If l IsNot Nothing Then
    '                tvvm2 = TreeViewVMRecursive(l)
    '                If tvvm2 IsNot Nothing Then
    '                    tvvm.Child.Add(tvvm2)
    '                End If
    '                'tvvm.Child.Add(TreeViewVMRecursive(l))
    '            End If
    '            '-------------------------------------------------------------
    '        Next
    '    End If
    '    Return tvvm
    'End Function

    '    'フレキシブルデザイン関係
    '    Public defaultWinHeight As System.Int32

    '    'データベース接続関係
    '    Dim scon As System.Data.SqlClient.SqlConnection
    '    Dim strn As System.Data.SqlClient.SqlTransaction
    '    Dim sda As System.Data.SqlClient.SqlDataAdapter
    '    Dim scmd As System.Data.SqlClient.SqlCommand
    '    Dim scb As System.Data.SqlClient.SqlCommandBuilder
    '    Dim ConStr As System.String
    '    Dim ds As System.Data.DataSet
    '    Dim dt As System.Data.DataTable
    '    Dim dr As System.Data.DataRow

    '    'セーブファイル関係
    '    Dim sr As System.IO.StreamReader
    '    Dim sw As System.IO.StreamWriter
    '    Private Const SHIFT_JIS = "SHIFT_JIS"


    '    'デリゲート関係
    '    Delegate Sub mySub()
    '    Delegate Sub myExSub(ByRef ex As System.Exception)
    '    Delegate Function myExFunc(ByVal ex As System.Exception)
    '    Dim ms As mySub
    '    Dim ms2 As mySub
    '    Dim mexs As myExSub
    '    Dim mexs2 As myExSub

    '    Private Const WINM = "MainWindow"
    '    Private Const WIN2 = "Window2"
    '    Private Const WIN3 = "Window3"

    '    Public Structure BData
    '        Public ServerName As System.String
    '        Public DBName As System.String
    '        Public FieldName As System.String
    '        Public SrcValue As System.String
    '        Public DistValue As System.String
    '        Public Query As System.String
    '        Public DTName As System.String
    '        Public TestDT As System.String
    '        Public AccessCk As System.Boolean
    '        Public SqlVer As System.String
    '        Public TableList() As System.String
    '    End Structure




    '    '    'ページ遷移
    '    '    Public Sub swichPages(ByRef OldWin As System.Windows.Window,
    '    '                          ByRef Btn As System.Windows.Controls.Button)


    '    '        Dim NewWin As System.Windows.Window

    '    '        Select Case OldWin.Title
    '    '            Case WINM
    '    '                If Btn.Name = "Next" Then
    '    '                    NewWin = New test2.Window2
    '    '                Else
    '    '                End If
    '    '            Case WIN2
    '    '                If Btn.Name = "Next" Then
    '    '                    NewWin = New test2.Window3
    '    '                Else
    '    '                    NewWin = New test2.MainWindow
    '    '                End If
    '    '            Case WIN3
    '    '                If Btn.Name = "Next" Then
    '    '                Else
    '    '                    NewWin = New test2.Window2
    '    '                End If
    '    '            Case Else
    '    '                NewWin = Nothing
    '    '        End Select

    '    '        NewWin.Show()

    '    '        OldWin.Close()
    '    '    End Sub


    '    '    'Private Sub Connection(ByVal f1 As System.Action,
    '    '    '                       ByVal f2 As System.Action(Of Exception),
    '    '    '                       ByVal f3 As System.Action,
    '    '    '                       ByVal f4 As System.Action(Of Exception))

    '    '    '    Try
    '    '    '        scon = New System.Data.SqlClient.SqlConnection(ConStr)
    '    '    '        scon.Open()
    '    '    '        strn = scon.BeginTransaction()

    '    '    '        'Connection Sucess
    '    '    '        Call f1()

    '    '    '    Catch ex As Exception
    '    '    '        'Connection Failed
    '    '    '        Call f2(ex)

    '    '    '        Try
    '    '    '            strn.Rollback()

    '    '    '        Catch ex2 As Exception
    '    '    '            'Rollback Failed
    '    '    '            Call f4(ex2)
    '    '    '        End Try

    '    '    '    Finally
    '    '    '        'Common Transaction
    '    '    '        Call f3()
    '    '    '        scon.Close()
    '    '    '        scon = Nothing
    '    '    '    End Try
    '    '    'End Sub

    '    Private Sub ConnectionEstablish()
    '        Try
    '            scon = New System.Data.SqlClient.SqlConnection(ConStr)
    '            scon.Open()
    '            strn = scon.BeginTransaction()
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Sub

    '    Private Sub ConnectionDispose()
    '        If scon IsNot Nothing Then
    '            scon.Close()
    '            scon.Dispose()
    '        End If
    '    End Sub

    '    Private Sub TransactionRollback()
    '        If strn IsNot Nothing Then
    '            strn.Rollback()
    '        End If
    '    End Sub

    '    Private Sub RollbackFailed(ByRef ex As System.Exception)
    '        MsgBox("Rollback Exception Type: " & ex.GetType().ToString)
    '        ' This catch block will handle any errors that may have occured
    '        ' on the server that would cause the rollback to fail, such as
    '        ' a closed connection
    '    End Sub

    '    Private Sub ConnectionFailed(ByRef ex As System.Exception)
    '        MsgBox("Commit Exception Type: " & ex.GetType().ToString)
    '    End Sub

    'Private Sub CommandEstablish(ByVal sql As SqlModel)
    '    scmd = New System.Data.SqlClient.SqlCommand
    '    scmd = scon.CreateCommand()
    '    scmd.CommandText = bd.Query
    '    scmd.CommandType = System.Data.CommandType.Text
    '    scmd.CommandTimeout = 30
    '    scmd.Transaction = strn
    'End Sub


    'Public Function AccessTest(ByVal sql As SqlModel) As SqlModel
    '    Dim ms3 As mySub
    '    Dim scmd As SqlCommand
    '    Dim scon As SqlConnection
    '    Dim strn As SqlTransaction

    '    Try
    '        '接続開始
    '        scon = New SqlConnection(sql.ConnectionText)
    '        scon.Open()
    '        strn = scon.BeginTransaction()

    '        'コマンド作成
    '        scmd = New System.Data.SqlClient.SqlCommand
    '        scmd = scon.CreateCommand()
    '        scmd.CommandText = sql.Query
    '        scmd.CommandType = System.Data.CommandType.Text
    '        scmd.CommandTimeout = 30
    '        scmd.Transaction = strn


    '        sql.ServerVersion = scon.ServerVersion

    '        '成功
    '        sql.AccessFlag = True
    '    Catch ex As Exception
    '        '失敗
    '        sql.AccessFlag = False

    '        Try
    '            If strn IsNot Nothing Then
    '                strn.Rollback()
    '            End If
    '        Catch ex2 As Exception
    '            '
    '        End Try
    '    Finally
    '        If scon IsNot Nothing Then
    '            scon.Close()
    '            scon.Dispose()
    '        End If
    '        AccessTest = sql
    '    End Try
    'End Function





    '    '    'データベースに存在するユーザテーブル一覧の取得
    '    '    'テストテーブル名の設定
    '    '    Public Function myTableList(ByVal bd As test2.testmodule.BData) As test2.testmodule.BData
    '    '        Dim cnt As System.Int32
    '    '        Dim ms3 As mySub
    '    '        Dim hitFlg As System.Boolean : hitFlg = Microsoft.VisualBasic.Constants.vbFalse
    '    '        Dim j As System.Int32

    '    '        ms = AddressOf ConnectionEstablish
    '    '        ms2 = AddressOf ConnectionDispose
    '    '        ms3 = AddressOf TransactionRollback
    '    '        mexs = AddressOf ConnectionFailed
    '    '        mexs2 = AddressOf RollbackFailed


    '    '        bd.Query = "SELECT * FROM SYS.OBJECTS WHERE TYPE = 'U'"
    '    '        ConStr = ConnectionString(bd)

    '    '        Try
    '    '            '接続...
    '    '            Call ms()

    '    '            'データ取得
    '    '            ds = test2.testmodule.myDataSet(bd)
    '    '            dt = ds.Tables(0)

    '    '            cnt = dt.Rows.Count

    '    '            'テーブル一覧
    '    '            ReDim bd.TableList(cnt - 1)
    '    '            For i = LBound(bd.TableList) To UBound(bd.TableList)
    '    '                bd.TableList(i) = dt.Rows(i)("name")
    '    '            Next


    '    '            'テストテーブル
    '    '            While Constants.vbTrue
    '    '                bd.TestDT = "TSTTBL" + j.ToString
    '    '                For i = LBound(bd.TableList) To UBound(bd.TableList)
    '    '                    If bd.TableList(i) = bd.TestDT Then
    '    '                        hitFlg = Constants.vbTrue
    '    '                        Exit For
    '    '                    End If
    '    '                Next

    '    '                If hitFlg Then
    '    '                    hitFlg = Constants.vbFalse
    '    '                Else
    '    '                    Exit While
    '    '                End If
    '    '            End While


    '    '        Catch ex As Exception
    '    '            'Failed ...
    '    '            bd.AccessCk = False
    '    '            Call mexs(ex)

    '    '            Try
    '    '                Call ms3()
    '    '            Catch ex2 As Exception
    '    '                Call mexs(ex2)
    '    '            End Try
    '    '        Finally
    '    '            myTableList = bd
    '    '            Call ms2()
    '    '        End Try
    '    '    End Function





    '    '    'クエリ単純実行
    '    '    Private Sub myQuery(ByVal bd As test2.testmodule.BData)
    '    '        mbds = AddressOf CommandEstablish
    '    '        Try
    '    '            Call mbds(bd)
    '    '            scmd.ExecuteNonQuery()
    '    '        Catch ex As Exception
    '    '            Throw ex
    '    '        Finally
    '    '            scmd.Dispose()
    '    '        End Try
    '    '    End Sub






    '    '    'クエリによるデータセット取得
    '    '    Private Function myDataSet(ByVal bd As test2.testmodule.BData) As System.Data.DataSet
    '    '        mbds = AddressOf CommandEstablish
    '    '        Try
    '    '            Call mbds(bd)
    '    '            sda = New System.Data.SqlClient.SqlDataAdapter(scmd)
    '    '            ds = New System.Data.DataSet
    '    '            sda.Fill(ds)
    '    '            myDataSet = ds

    '    '        Catch ex As Exception
    '    '            Throw ex
    '    '        Finally
    '    '            sda.Dispose()
    '    '            ds.Dispose()
    '    '            scmd.Dispose()
    '    '        End Try
    '    '        '---------------------------------------------------------------------------
    '    '    End Function



    '    '    'テストテーブルとやり取りするタイプ
    '    '    Public Sub TestTableDelete(ByVal bd As test2.testmodule.BData)
    '    '        ms = AddressOf ConnectionEstablish
    '    '        ms2 = AddressOf ConnectionDispose
    '    '        mexs = AddressOf RollbackFailed

    '    '        ConStr = ConnectionString(bd)

    '    '        Try
    '    '            Call ms()

    '    '            'テストテーブル削除
    '    '            bd.Query = $"IF OBJECT_ID('{bd.TestDT}', 'U') IS NOT NULL BEGIN DROP TABLE {bd.TestDT} END"
    '    '            'If bd.SqlVer > 2015 Then
    '    '            '    bd.Query = $"IF EXISTS DROP TABLE {bd.TestDT}"
    '    '            'End If
    '    '            Call myQuery(bd)

    '    '            '成功
    '    '            strn.Commit()
    '    '        Catch ex As Exception
    '    '            '失敗
    '    '            Try
    '    '                strn.Rollback()
    '    '            Catch ex2 As Exception
    '    '                Call mexs(ex2)
    '    '            End Try
    '    '            MsgBox("作業テーブルの削除に失敗しました。" _
    '    '                & vbCrLf & "以下の名前で作業テーブルが残る可能性があります。" _
    '    '                & vbCrLf & bd.TestDT)
    '    '        Finally
    '    '            Call ms2()
    '    '        End Try
    '    '    End Sub


    '    '    'テストテーブルとやり取りするタイプ
    '    '    Public Function Main2(ByVal bd As test2.testmodule.BData) As System.String

    '    '        Dim sourceValue As System.Object
    '    '        Dim distinationvalue As System.Object

    '    '        Dim mexf As test2.testmodule.myExFunc

    '    '        ms = AddressOf ConnectionEstablish
    '    '        ms2 = AddressOf ConnectionDispose
    '    '        mexf = Function(ex)
    '    '                   Return ex.GetType().ToString() & " " & ex.Message
    '    '               End Function
    '    '        mexs = AddressOf RollbackFailed

    '    '        '戻り値テキスト
    '    '        Dim result As System.String : result = Constants.vbNullString
    '    '        Dim msg As System.String : msg = Constants.vbNullString
    '    '        Dim errmsg As System.String : errmsg = Constants.vbNullString

    '    '        ConStr = ConnectionString(bd)

    '    '        Try
    '    '            Call ms()

    '    '            sourceValue = "'" & bd.SrcValue & "'"
    '    '            distinationvalue = "'" & bd.DistValue & "'"

    '    '            'テストテーブル削除
    '    '            bd.Query = $"IF OBJECT_ID('{bd.TestDT}', 'U') IS NOT NULL BEGIN DROP TABLE {bd.TestDT} END"
    '    '            'If bd.SqlVer > 2015 Then
    '    '            '    bd.Query = $"IF EXISTS DROP TABLE {bd.TestDT}"
    '    '            'End If
    '    '            Call myQuery(bd)

    '    '            'テストテーブル作成
    '    '            bd.Query = $"SELECT * INTO {bd.TestDT} FROM {bd.DTName} WHERE 1 <> 1"
    '    '            Call myQuery(bd)

    '    '            'テストテーブル複製
    '    '            bd.Query = $"INSERT INTO {bd.TestDT} SELECT * FROM {bd.DTName} WHERE {bd.FieldName} = {sourceValue}"
    '    '            Call myQuery(bd)

    '    '            'テストテーブル更新
    '    '            bd.Query = $"UPDATE {bd.TestDT} SET {bd.FieldName} = {distinationvalue} WHERE {bd.FieldName} = {sourceValue}"
    '    '            Call myQuery(bd)

    '    '            'テストテーブルから挿入
    '    '            bd.Query = $"INSERT INTO {bd.DTName} SELECT * FROM {bd.TestDT} WHERE {bd.FieldName} = {bd.DistValue}"
    '    '            Call myQuery(bd)

    '    '            '成功
    '    '            strn.Commit()
    '    '            result = "O K"
    '    '            msg = bd.SrcValue & " --> " & bd.DistValue
    '    '        Catch ex As Exception
    '    '            '失敗
    '    '            result = "ERR"
    '    '            errmsg = mexf(ex)

    '    '            Try
    '    '                strn.Rollback()
    '    '            Catch ex2 As Exception
    '    '                Call mexs(ex2)
    '    '            End Try
    '    '        Finally
    '    '            Main2 = "(" & result & ")" & "[" & bd.DTName & " / " _
    '    '                & bd.FieldName & "] " & msg & " " & errmsg
    '    '            Call ms2()
    '    '        End Try
    '    '        '---------------------------------------------------------------------------
    '    '    End Function



    '    '    ''メモリー内でやり取りするタイプ
    '    '    'Public Function Main(ByVal bd As test2.testmodule.BData) As System.String

    '    '    '    'DateTime型対応
    '    '    '    Dim sourceValue As System.Object
    '    '    '    Dim distinationvalue As System.Object

    '    '    '    '戻り値テキスト
    '    '    '    Dim result As System.String : result = Constants.vbNullString
    '    '    '    Dim msg As System.String : msg = Constants.vbNullString
    '    '    '    Dim errmsg As System.String : errmsg = Constants.vbNullString
    '    '    '    Dim err2msg As System.String : err2msg = Constants.vbNullString

    '    '    '    '前後にカンマを付加
    '    '    '    '空白を含むフィールド値に対応するため
    '    '    '    sourceValue = "'" & bd.SrcValue & "'"

    '    '    '    '抽出クエリ文
    '    '    '    bd.Query = $"SELECT * FROM {bd.DTName} WHERE {bd.FieldName} = {sourceValue}"

    '    '    '    '接続文字列
    '    '    '    ConStr = ConnectionString(bd)

    '    '    '    '---------------------------------------------------------------------------
    '    '    '    Try
    '    '    '        scon = New System.Data.SqlClient.SqlConnection(ConStr)
    '    '    '        scon.Open()
    '    '    '        strn = scon.BeginTransaction()

    '    '    '        ds = test2.testmodule.myDataSet(bd)

    '    '    '        dt = ds.Tables(0)


    '    '    '        '書き方が格好良くないが、頭の１件を見てデータ型を判定する
    '    '    '        Select Case dt.Rows(1)(bd.FieldName).GetType().Name
    '    '    '            Case "DateTime"
    '    '    '                sourceValue = Convert.ToDateTime(bd.SrcValue)
    '    '    '                distinationvalue = Convert.ToDateTime(bd.DistValue)
    '    '    '            Case Else
    '    '    '                sourceValue = bd.SrcValue
    '    '    '                distinationvalue = bd.DistValue
    '    '    '        End Select

    '    '    '        '色々考えたが、無理だった・・・
    '    '    '        'dt2 = dt.AsEnumerable().Where(
    '    '    '        '    Function(r)
    '    '    '        '        Return r(bd.FieldName) = bd.SrcValue
    '    '    '        '    End Function
    '    '    '        ').Select(
    '    '    '        '    Function(r)
    '    '    '        '        r(bd.FieldName) = bd.DistValue
    '    '    '        '        Return r
    '    '    '        '    End Function
    '    '    '        ').CopyToDataTable()


    '    '    '        For i As Integer = 0 To dt.Rows.Count
    '    '    '            If dt.Rows(i)(bd.FieldName) = sourceValue Then
    '    '    '                dr = dt.NewRow
    '    '    '                For Each c In dt.Columns
    '    '    '                    If c.ToString = bd.FieldName Then
    '    '    '                        dr(c.ToString) = distinationvalue
    '    '    '                    Else
    '    '    '                        dr(c.ToString) = dt.Rows(i)(c.ToString)
    '    '    '                    End If
    '    '    '                Next
    '    '    '                dt.Rows.Add(dr)
    '    '    '            End If
    '    '    '        Next

    '    '    '        scb = New System.Data.SqlClient.SqlCommandBuilder(sda)
    '    '    '        sda.Update(ds, dt.TableName)

    '    '    '        strn.Commit()

    '    '    '        'result OK
    '    '    '        result = "O K"
    '    '    '        msg = bd.SrcValue & " --> " & bd.DistValue

    '    '    '    Catch ex As Exception
    '    '    '        'result ERR
    '    '    '        result = "ERR"
    '    '    '        errmsg = ex.GetType().ToString() & " " & ex.Message

    '    '    '        Try
    '    '    '            strn.Rollback()
    '    '    '        Catch ex2 As Exception
    '    '    '            errmsg = ex2.GetType().ToString() & " " & ex2.Message
    '    '    '            'MsgBox("Rollback Exception Type: " & ex2.GetType().ToString)
    '    '    '            ' This catch block will handle any errors that may have occured
    '    '    '            ' on the server that would cause the rollback to fail, such as
    '    '    '            ' a closed connection
    '    '    '        End Try

    '    '    '    Finally
    '    '    '        Main = "(" & result & ")" & "[" & bd.DTName & " / " _
    '    '    '            & bd.FieldName & "] " & msg & " " & errmsg & " " & err2msg
    '    '    '        scon.Close()
    '    '    '        scon = Nothing
    '    '    '    End Try
    '    '    '    '---------------------------------------------------------------------------
    '    '    'End Function




    '    '    Private ReadOnly Property ConnectionString(ByVal bd As test2.testmodule.BData) As System.String
    '    '        Get
    '    '            Return "Data Source=" & bd.ServerName & ";" _
    '    '                & "Initial Catalog=" & bd.DBName & ";" _
    '    '                & "Integrated Security=True;" _
    '    '                & "Connect Timeout=05;"
    '    '        End Get
    '    '    End Property




    '    '    Private ReadOnly Property saveDir() As System.String
    '    '        Get
    '    '            Return System.Environment.GetEnvironmentVariable("USERPROFILE") _
    '    '                & "\test"
    '    '        End Get
    '    '    End Property




    '    '    Private ReadOnly Property iniFile() As System.String
    '    '        Get
    '    '            Return saveDir & "\" & "iniFile"
    '    '        End Get
    '    '    End Property




    '    '    Private ReadOnly Property saveFile() As System.String
    '    '        Get
    '    '            saveFile = saveDir & "\" & "saveFile"
    '    '        End Get
    '    '    End Property


    '    '    Private Function InspectBdata(ByVal bd As test2.testmodule.BData) As test2.testmodule.BData
    '    '        'この関数は例外を出すことはないはず・・・
    '    '        If String.IsNullOrEmpty(bd.ServerName) Then
    '    '            bd.ServerName = "Server Name"
    '    '        End If
    '    '        If String.IsNullOrEmpty(bd.DBName) Then
    '    '            bd.DBName = "DataBase Name"
    '    '        End If
    '    '        If String.IsNullOrEmpty(bd.FieldName) Then
    '    '            bd.FieldName = "Field Name"
    '    '        End If
    '    '        If String.IsNullOrEmpty(bd.SrcValue) Then
    '    '            bd.SrcValue = "Source Value"
    '    '        End If
    '    '        If String.IsNullOrEmpty(bd.DistValue) Then
    '    '            bd.DistValue = "Distination Value"
    '    '        End If
    '    '        If String.IsNullOrEmpty(bd.TestDT) Then
    '    '            bd.TestDT = "TESTDB"
    '    '        End If
    '    '        'If String.IsNullOrEmpty(bd.SqlVer) Then
    '    '        '    bd.SqlVer = DEFAULT_SQL_VER
    '    '        'End If
    '    '        'If IsNothing(bd.TableList) Then
    '    '        '    bd.TableList("name") = Constants.vbNullString
    '    '        '    bd.TableList("stat") = Constants.vbNullString
    '    '        'End If

    '    '        InspectBdata = bd
    '    '    End Function


    '    '    Public Sub Save(ByVal bd As test2.testmodule.BData)
    '    '        Dim myJson As System.String
    '    '        Dim i As System.Int32 : i = 0

    '    '        '例外はいずれもセーブせず終了
    '    '        Try
    '    '            i = test2.testmodule.SaveCheck
    '    '            bd = test2.testmodule.InspectBdata(bd)

    '    '            myJson = JsonConvert.SerializeObject(bd)
    '    '            sw = New System.IO.StreamWriter(
    '    '                test2.testmodule.saveFile, False, System.Text.Encoding.GetEncoding(SHIFT_JIS))
    '    '            sw.WriteLine(myJson)
    '    '        Catch ex As Exception
    '    '            bd = test2.testmodule.InitializeBData(bd)
    '    '        Finally
    '    '            If sw IsNot Nothing Then
    '    '                sw.Close()
    '    '                sw.Dispose()
    '    '            End If
    '    '        End Try
    '    '    End Sub

    '    '    Private Function SaveCheck() As System.Int32
    '    '        Try
    '    '            If Not System.IO.Directory.Exists(test2.testmodule.saveDir) Then
    '    '                System.IO.Directory.CreateDirectory(test2.testmodule.saveDir)
    '    '                System.IO.File.Create(test2.testmodule.iniFile)
    '    '            End If

    '    '            '既に同名フォルダーが存在する場合、作成しない
    '    '            If Not System.IO.File.Exists(test2.testmodule.iniFile) Then
    '    '                Throw New Exception
    '    '            End If

    '    '            'saveFileが存在しない(初回)
    '    '            If Not System.IO.File.Exists(test2.testmodule.saveFile) Then
    '    '                System.IO.File.Create(test2.testmodule.saveFile)
    '    '                SaveCheck = 2
    '    '                Exit Function
    '    '            End If

    '    '            '正常
    '    '            SaveCheck = 1
    '    '        Catch ex As Exception
    '    '            Throw ex
    '    '        End Try
    '    '    End Function

    '    '    Private Function InitializeBData(ByVal bd As test2.testmodule.BData) As test2.testmodule.BData
    '    '        bd.ServerName = "Server Name"
    '    '        bd.DBName = "DataBase Name"
    '    '        bd.FieldName = "Field Name"
    '    '        bd.SrcValue = "Source Value"
    '    '        bd.DistValue = "Distination Value"
    '    '        bd.TestDT = "TESTDB"
    '    '        'bd.SqlVer = DEFAULT_SQL_VER
    '    '        'If Not IsNothing(bd.TableList) Then
    '    '        '    bd.TableList = Nothing
    '    '        'End If
    '    '        'bd.TableList("name") = Constants.vbNullString
    '    '        'bd.TableList("stat") = Constants.vbNullString

    '    '        InitializeBData = bd
    '    '    End Function


    '    '    Public Function Load(ByVal bd As test2.testmodule.BData) As test2.testmodule.BData
    '    '        Dim myJson As System.String
    '    '        Dim obj As test2.testmodule.BData

    '    '        'いずれの例外も初期化
    '    '        Try
    '    '            If SaveCheck() = 2 Then
    '    '                Throw New Exception
    '    '            End If

    '    '            sr = New System.IO.StreamReader(
    '    '            test2.testmodule.saveFile, System.Text.Encoding.GetEncoding(SHIFT_JIS))

    '    '            myJson = sr.ReadToEnd()
    '    '            obj = JsonConvert.DeserializeObject(Of test2.testmodule.BData)(myJson)

    '    '            'エラーが発生し得る
    '    '            bd.ServerName = obj.ServerName
    '    '            bd.DBName = obj.DBName
    '    '            bd.FieldName = obj.FieldName
    '    '            bd.SrcValue = obj.SrcValue
    '    '            bd.DistValue = obj.DistValue
    '    '            bd.TestDT = obj.TestDT
    '    '            'bd.SqlVer = obj.SqlVer

    '    '        Catch ex As Exception

    '    '            bd = InitializeBData(bd)
    '    '        Finally
    '    '            Load = bd
    '    '            If sr IsNot Nothing Then
    '    '                sr.Close()
    '    '                sr = Nothing
    '    '            End If
    '    '            obj = Nothing
    '    '        End Try
    '    '    End Function


    '    '    Public Sub RequirementCheck()
    '    '        Try
    '    '            'ThreadPool And Rambda
    '    '            ThreadPool.QueueUserWorkItem(
    '    '            Sub()
    '    '            End Sub, Nothing)

    '    '            '
    '    '            String.IsNullOrEmpty("")

    '    '            '
    '    '            System.Environment.GetEnvironmentVariable("USERPROFILE")

    '    '            '
    '    '        Catch ex As Exception

    '    '        End Try
    '    '    End Sub

    '    '    Public Function changeFontSize(txtSize As System.Int32, winHeight As System.Int32) As System.Int32
    '    '        changeFontSize = (txtSize / test2.testmodule.defaultWinHeight) * winHeight
    '    '    End Function
End Module
