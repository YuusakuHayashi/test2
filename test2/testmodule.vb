Imports Newtonsoft.Json

Module testmodule
    'フレキシブルデザイン関係
    Public defaultWinHeight As System.Int32

    'データベース接続関係
    Dim scon As System.Data.SqlClient.SqlConnection
    Dim strn As System.Data.SqlClient.SqlTransaction
    Dim sda As System.Data.SqlClient.SqlDataAdapter
    Dim scmd As System.Data.SqlClient.SqlCommand
    Dim scb As System.Data.SqlClient.SqlCommandBuilder
    Dim ConStr As System.String
    Dim ds As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dr As System.Data.DataRow

    'セーブファイル関係
    Dim sr As System.IO.StreamReader
    Dim sw As System.IO.StreamWriter


    Delegate Sub connectionDelegator(ByVal r1 As Action,
                                     ByVal r2 As Action(Of Exception),
                                     ByVal r3 As Action,
                                     ByVal r4 As Action(Of Exception))



    Public Structure BData
        Public ServerName As System.String
        Public DBName As System.String
        Public FieldName As System.String
        Public SrcValue As System.String
        Public DistValue As System.String
        Public Query As System.String
        Public DTName As System.String
    End Structure



    'ページ遷移
    Public Sub swichPages(ByRef OldWin As System.Windows.Window,
                          ByRef Btn As System.Windows.Controls.Button)


        Dim NewWin As System.Windows.Window
        Const WINM = "MainWindow"
        Const WIN2 = "Window2"
        Const WIN3 = "Window3"

        Select Case OldWin.Title
            Case WINM
                If Btn.Name = "Next" Then
                    NewWin = New test2.Window2
                Else
                End If
            Case WIN2
                If Btn.Name = "Next" Then
                    NewWin = New test2.Window3
                Else
                    NewWin = New test2.MainWindow
                End If
            Case WIN3
                If Btn.Name = "Next" Then
                Else
                    NewWin = New test2.Window2
                End If
            Case Else
                NewWin = Nothing
        End Select

        NewWin.Show()

        OldWin.Close()
    End Sub


    'Private Sub Connection(ByVal f1 As System.Action,
    '                       ByVal f2 As System.Action(Of Exception),
    '                       ByVal f3 As System.Action,
    '                       ByVal f4 As System.Action(Of Exception))

    '    Try
    '        scon = New System.Data.SqlClient.SqlConnection(ConStr)
    '        scon.Open()
    '        strn = scon.BeginTransaction()

    '        'Connection Sucess
    '        Call f1()

    '    Catch ex As Exception
    '        'Connection Failed
    '        Call f2(ex)

    '        Try
    '            strn.Rollback()

    '        Catch ex2 As Exception
    '            'Rollback Failed
    '            Call f4(ex2)
    '        End Try

    '    Finally
    '        'Common Transaction
    '        Call f3()
    '        scon.Close()
    '        scon = Nothing
    '    End Try
    'End Sub




    Public Function AccessTest(ByVal bd As test2.testmodule.BData) As System.Boolean
        Dim rtn As System.Boolean
        Dim trnFlg As System.Boolean

        'Transactionが実行したかどうかのフラグ
        trnFlg = Constants.vbFalse

        'データベースの接続文字列
        ConStr = ConnectionString(bd)

        '' Connection Sucess
        'Dim rtn1 As System.Action = Sub()
        '                                rtn = Constants.vbTrue
        '                            End Sub

        '' Connection Failed
        'Dim rtn2 = Sub(ByRef ex)
        '               rtn = Constants.vbFalse
        '               MsgBox("Commit Exception Type: " & ex.GetType().ToString)
        '           End Sub

        '' Common Transaction
        'Dim rtn3 As System.Action = Sub()
        '                                testAccess = rtn
        '                            End Sub

        '' Rollback Failed
        'Dim rtn4 = Sub(ByRef ex2)
        '               MsgBox("Rollback Exception Type: " & ex2.GetType().ToString)
        '               ' This catch block will handle any errors that may have occured
        '               ' on the server that would cause the rollback to fail, such as
        '               ' a closed connection
        '           End Sub

        ''Dim delg As New connectionDelegator(AddressOf Connection)
        'Call Connection(rtn1, rtn2, rtn3, rtn4)

        '---------------------------------------------------------------------------
        Try
            scon = New System.Data.SqlClient.SqlConnection(ConStr)
            scon.Open()
            strn = scon.BeginTransaction()
            trnFlg = Constants.vbTrue

            'Sucess ...
            rtn = Constants.vbTrue

        Catch ex As Exception

            'Failed ...
            rtn = Constants.vbFalse

            Try
                If trnFlg Then
                    strn.Rollback()
                End If
            Catch ex2 As Exception
                'MsgBox("Rollback Exception Type: " & ex2.GetType().ToString)
                ' This catch block will handle any errors that may have occured
                ' on the server that would cause the rollback to fail, such as
                ' a closed connection
            End Try

        Finally
            scon.Close()
            AccessTest = rtn
            scon = Nothing
        End Try
        '---------------------------------------------------------------------------
    End Function



    'データベースに存在するユーザテーブル一覧の取得
    Public Function myTableList(ByVal bd As test2.testmodule.BData) As System.String()
        Dim Tbl() As System.String
        Dim cnt As System.Int32

        bd.Query = "SELECT * FROM SYS.OBJECTS WHERE TYPE = 'U'"

        ConStr = ConnectionString(bd)

        '---------------------------------------------------------------------------
        Try
            '接続...
            scon = New System.Data.SqlClient.SqlConnection(ConStr)
            scon.Open()
            strn = scon.BeginTransaction()

            'データ取得
            ds = test2.testmodule.myDataSet(bd)
            dt = ds.Tables(0)

            cnt = dt.Rows.Count

            ReDim Tbl(cnt - 1)
            For i = LBound(Tbl) To UBound(Tbl)
                Tbl(i) = dt.Rows(i)("name")
            Next

            'Success ...
            myTableList = Tbl

        Catch ex As Exception
            'Failed ...
            MsgBox("Commit Exception Type: " & ex.GetType().ToString)

            Try
                strn.Rollback()
            Catch ex2 As Exception
                MsgBox("Rollback Exception Type: " & ex2.GetType().ToString)
                ' This catch block will handle any errors that may have occured
                ' on the server that would cause the rollback to fail, such as
                ' a closed connection
            End Try

        Finally
            scon.Close()
            scon = Nothing
        End Try
        '---------------------------------------------------------------------------
    End Function




    'クエリによるデータセット取得
    Private Function myDataSet(ByVal bd As test2.testmodule.BData) As System.Data.DataSet
        '---------------------------------------------------------------------------
        Try
            scmd = New System.Data.SqlClient.SqlCommand
            scmd = scon.CreateCommand()
            scmd.CommandText = bd.Query
            scmd.CommandType = System.Data.CommandType.Text
            scmd.CommandTimeout = 30
            scmd.Transaction = strn

            sda = New System.Data.SqlClient.SqlDataAdapter(scmd)
            ds = New System.Data.DataSet
            sda.Fill(ds)
            myDataSet = ds

        Catch ex As Exception
            Throw ex
        Finally
            scmd = Nothing
        End Try
        '---------------------------------------------------------------------------
    End Function




    Public Function Main(ByVal bd As test2.testmodule.BData) As System.String

        'DateTime型対応
        Dim sourceValue As System.Object
        Dim distinationvalue As System.Object

        '戻り値テキスト
        Dim result As System.String : result = Constants.vbNullString
        Dim msg As System.String : msg = Constants.vbNullString
        Dim errmsg As System.String : errmsg = Constants.vbNullString
        Dim err2msg As System.String : err2msg = Constants.vbNullString


        '前後にカンマを付加
        '空白を含むフィールド値に対応するため
        sourceValue = "'" & bd.SrcValue & "'"

        '抽出クエリ文
        bd.Query = $"SELECT * FROM {bd.DTName} WHERE {bd.FieldName} = {sourceValue}"

        '接続文字列
        ConStr = ConnectionString(bd)

        '---------------------------------------------------------------------------
        Try
            scon = New System.Data.SqlClient.SqlConnection(ConStr)
            scon.Open()
            strn = scon.BeginTransaction()

            ds = test2.testmodule.myDataSet(bd)

            dt = ds.Tables(0)

            '書き方が格好良くないが、頭の１件を見てデータ型を判定する
            Select Case dt.Rows(1)(bd.FieldName).GetType().Name
                Case "DateTime"
                    sourceValue = Convert.ToDateTime(bd.SrcValue)
                    distinationvalue = Convert.ToDateTime(bd.DistValue)
                Case Else
                    sourceValue = bd.SrcValue
                    distinationvalue = bd.DistValue
            End Select

            For i As Integer = 0 To dt.Rows.Count - 1
                If dt.Rows(i)(bd.FieldName) = sourceValue Then
                    dr = dt.NewRow
                    For Each c In dt.Columns
                        If c.ToString = bd.FieldName Then
                            dr(c.ToString) = distinationvalue
                        Else
                            dr(c.ToString) = dt.Rows(i)(c.ToString)
                        End If
                    Next
                    dt.Rows.Add(dr)
                End If
            Next

            scb = New System.Data.SqlClient.SqlCommandBuilder(sda)
            sda.Update(ds, dt.TableName)

            strn.Commit()

            'result OK
            result = "O K"
            msg = bd.SrcValue & " --> " & bd.DistValue

        Catch ex As Exception
            'result ERR
            result = "ERR"
            errmsg = ex.GetType().ToString() & " " & ex.Message

            Try
                strn.Rollback()
            Catch ex2 As Exception
                errmsg = ex2.GetType().ToString() & " " & ex2.Message
                'MsgBox("Rollback Exception Type: " & ex2.GetType().ToString)
                ' This catch block will handle any errors that may have occured
                ' on the server that would cause the rollback to fail, such as
                ' a closed connection
            End Try

        Finally
            Main = "(" & result & ")" & "[" & bd.DTName & " / " _
                & bd.FieldName & "] " & msg & " " & errmsg & " " & err2msg
            scon.Close()
            scon = Nothing
        End Try
        '---------------------------------------------------------------------------
    End Function




    Private ReadOnly Property ConnectionString(ByVal bd As test2.testmodule.BData) As System.String
        Get
            Return "Data Source=" & bd.ServerName & ";" _
                & "Initial Catalog=" & bd.DBName & ";" _
                & "Integrated Security=True;" _
                & "Connect Timeout=05;"
        End Get
    End Property




    Private ReadOnly Property saveDir() As System.String
        Get
            Return System.Environment.GetEnvironmentVariable("USERPROFILE") _
                & "\test"
        End Get
    End Property




    Private ReadOnly Property iniFile() As System.String
        Get
            Return saveDir & "\" & "iniFile"
        End Get
    End Property




    Private ReadOnly Property saveFile() As System.String
        Get
            saveFile = saveDir & "\" & "saveFile"
        End Get
    End Property




    Public Sub Save(ByVal bd As test2.testmodule.BData)
        Dim myJson As System.String

        If test2.testmodule.SaveCheck = 0 Then
            Exit Sub
        End If

        If bd.ServerName = Constants.vbNullString Then
            bd.ServerName = "Server Name"
        End If
        If bd.DBName = Constants.vbNullString Then
            bd.DBName = "DataBase Name"
        End If
        If bd.FieldName = Constants.vbNullString Then
            bd.FieldName = "Field Name"
        End If
        If bd.SrcValue = Constants.vbNullString Then
            bd.SrcValue = "Source Value"
        End If
        If bd.DistValue = Constants.vbNullString Then
            bd.DistValue = "Distination Value"
        End If

        myJson = JsonConvert.SerializeObject(bd)

        sw = New System.IO.StreamWriter(
            test2.testmodule.saveFile, False, System.Text.Encoding.GetEncoding("SHIFT_JIS"))

        sw.WriteLine(myJson)

        sw.Close()

        sw = Nothing
    End Sub

    Private Function SaveCheck() As System.Int32
        If Not System.IO.Directory.Exists(test2.testmodule.saveDir) Then
            System.IO.Directory.CreateDirectory(test2.testmodule.saveDir)
            System.IO.File.Create(test2.testmodule.iniFile)
        End If

        '既に同名フォルダーが存在する場合、作成しない
        If Not System.IO.File.Exists(test2.testmodule.iniFile) Then
            SaveCheck = 0
            Exit Function
        End If

        'saveFileが存在しない
        If Not System.IO.File.Exists(test2.testmodule.saveFile) Then
            System.IO.File.Create(test2.testmodule.saveFile)
            SaveCheck = 2
            Exit Function
        End If

        '正常
        SaveCheck = 1
    End Function

    Public Function Load(ByVal bd As test2.testmodule.BData) As test2.testmodule.BData
        Dim i As System.Int32 : i = 0
        Dim myJson As System.String
        Dim Obj As test2.testmodule.BData

        Load = bd
        If SaveCheck() = (0 Or 2) Then
            Exit Function
        End If

        sr = New System.IO.StreamReader(
            test2.testmodule.saveFile, System.Text.Encoding.GetEncoding("SHIFT_JIS"))

        myJson = sr.ReadToEnd()
        Obj = JsonConvert.DeserializeObject(Of test2.testmodule.BData)(myJson)

        bd.ServerName = Obj.ServerName
        bd.DBName = Obj.DBName
        bd.FieldName = Obj.FieldName
        bd.SrcValue = Obj.SrcValue
        bd.DistValue = Obj.DistValue

        Load = bd

        sr.Close()

        sr = Nothing
        Obj = Nothing
    End Function

    Public Function changeFontSize(txtSize As System.Int32, winHeight As System.Int32) As System.Int32
        changeFontSize = (txtSize / test2.testmodule.defaultWinHeight) * winHeight
    End Function
End Module
