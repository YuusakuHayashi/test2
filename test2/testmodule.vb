Module testmodule
    'swithPage ----------------------------------------------------
    Private Const winm = "test2.MainWindow"
    Private Const win2 = "test2.Window2"
    Private Const win3 = "test2.Window3"

    Private Const nx = "次へ"
    Private Const bk = "戻る"

    Private win As System.Windows.Window
    Public defaultWinHeight As System.Int32

    '--------------------------------------------------------------

    'database manupirate --------------------------------------
    Dim Con As System.Data.SqlClient.SqlConnection
    Dim Trn As System.Data.SqlClient.SqlTransaction
    Dim sda As System.Data.SqlClient.SqlDataAdapter
    Dim sc As System.Data.SqlClient.SqlCommand
    Dim scb As System.Data.SqlClient.SqlCommandBuilder
    Dim conStr As System.String
    Dim ds As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dr As System.Data.DataRow
    Dim trnFlg As System.Boolean
    '--------------------------------------------------------------

    'savefile -------------------------------------------------
    Dim sr As System.IO.StreamReader
    Dim sw As System.IO.StreamWriter
    '--------------------------------------------------------------

    Delegate Sub connectionDelegator(ByVal r1 As Action,
                                     ByVal r2 As Action(Of Exception),
                                     ByVal r3 As Action,
                                     ByVal r4 As Action(Of Exception))


    Public Structure BData
        Public serverName As System.String
        Public databaseName As System.String
        Public fieldName As System.String
        Public sourceValue As System.String
        Public distinationValue As System.String
        Public query As System.String
        Public datatableName As System.String
    End Structure




    Public Sub swichPages(ByRef w, ByRef s)
        Dim flg As System.Boolean : flg = Constants.vbFalse
        Select Case w.ToString
            Case winm
                If s.Content = nx Then
                    win = New test2.Window2
                Else
                    '
                End If
            Case win2
                If s.Content = nx Then
                    win = New test2.Window3()
                Else
                    win = New test2.MainWindow
                    flg = Constants.vbTrue
                End If
            Case win3
                If s.Content = nx Then
                    '
                Else
                    win = New test2.Window2
                End If
        End Select

        'Switch Window

        If flg Then
            win.Show()
        Else
            win.Show()
        End If

        'Close Window
        w.Close()
    End Sub


    Private Sub Connection(ByVal f1 As System.Action,
                           ByVal f2 As System.Action(Of Exception),
                           ByVal f3 As System.Action,
                           ByVal f4 As System.Action(Of Exception))

        Try
            Con = New System.Data.SqlClient.SqlConnection(conStr)
            Con.Open()
            Trn = Con.BeginTransaction()

            'Connection Sucess
            Call f1()

        Catch ex As Exception
            'Connection Failed
            Call f2(ex)

            Try
                Trn.Rollback()

            Catch ex2 As Exception
                'Rollback Failed
                Call f4(ex2)
            End Try

        Finally
            'Common Transaction
            Call f3()
            Con.Close()
            Con = Nothing
        End Try
    End Sub

    Public Function testAccess(ByVal bd As test2.testmodule.BData) As System.Boolean
        Dim rtn As Boolean

        trnFlg = False
        conStr = connectionString(bd)

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
            Con = New System.Data.SqlClient.SqlConnection(conStr)
            Con.Open()
            Trn = Con.BeginTransaction()
            trnFlg = True

            'sucess
            rtn = Constants.vbTrue

        Catch ex As Exception
            'failed
            rtn = Constants.vbFalse
            'MsgBox("Commit Exception Type: " & ex.GetType().ToString)

            Try
                If trnFlg Then
                    Trn.Rollback()
                End If
            Catch ex2 As Exception
                'MsgBox("Rollback Exception Type: " & ex2.GetType().ToString)
                ' This catch block will handle any errors that may have occured
                ' on the server that would cause the rollback to fail, such as
                ' a closed connection
            End Try

        Finally
            Con.Close()
            testAccess = rtn
            Con = Nothing
        End Try
        '---------------------------------------------------------------------------
    End Function

    Public Function TableList(ByVal bd As test2.testmodule.BData) As System.String()
        Dim s() As System.String
        Dim cnt As System.Int32

        bd.query = "SELECT * FROM SYS.OBJECTS WHERE TYPE = 'U'"

        conStr = connectionString(bd)

        '---------------------------------------------------------------------------
        Try
            Con = New System.Data.SqlClient.SqlConnection(conStr)
            Con.Open()
            Trn = Con.BeginTransaction()

            ds = test2.testmodule.getDataSet(bd)
            dt = ds.Tables(0)

            cnt = dt.Rows.Count

            ReDim s(cnt - 1)
            For i = LBound(s) To UBound(s)
                s(i) = dt.Rows(i)("name")
            Next
            TableList = s

        Catch ex As Exception
            'failed
            MsgBox("Commit Exception Type: " & ex.GetType().ToString)

            Try
                Trn.Rollback()
            Catch ex2 As Exception
                MsgBox("Rollback Exception Type: " & ex2.GetType().ToString)
                ' This catch block will handle any errors that may have occured
                ' on the server that would cause the rollback to fail, such as
                ' a closed connection
            End Try
        Finally
            Con.Close()
            Con = Nothing
        End Try
        '---------------------------------------------------------------------------
    End Function


    'Get & Return Data Set
    Private Function getDataSet(ByVal bd As test2.testmodule.BData) As System.Data.DataSet
        '---------------------------------------------------------------------------
        Try
            sc = New System.Data.SqlClient.SqlCommand
            sc = Con.CreateCommand()
            sc.CommandText = bd.query
            sc.CommandType = System.Data.CommandType.Text
            sc.CommandTimeout = 30
            sc.Transaction = Trn

            sda = New System.Data.SqlClient.SqlDataAdapter(sc)
            ds = New System.Data.DataSet
            sda.Fill(ds)
            getDataSet = ds

        Catch ex As Exception
            Throw ex
        Finally
            sc = Nothing
        End Try
        '---------------------------------------------------------------------------
    End Function


    Public Function Main(ByVal bd As test2.testmodule.BData) As System.String
        Dim lvl As System.Int32 : lvl = 0

        Dim sourceValue As System.Object
        Dim distinationvalue As System.Object

        Dim result As System.String : result = Constants.vbNullString
        Dim msg As System.String : msg = Constants.vbNullString
        Dim errmsg As System.String : errmsg = Constants.vbNullString
        Dim err2msg As System.String : err2msg = Constants.vbNullString


        'DateTime型のフィールドに対応するため、前後にカンマを付加
        sourceValue = "'" & bd.sourceValue & "'"

        bd.query = $"SELECT * FROM {bd.datatableName} WHERE {bd.fieldName} = {sourceValue}"

        conStr = connectionString(bd)

        '---------------------------------------------------------------------------
        Try
            Con = New System.Data.SqlClient.SqlConnection(conStr)
            Con.Open()
            Trn = Con.BeginTransaction()

            ds = test2.testmodule.getDataSet(bd)

            dt = ds.Tables(0)

            '書き方が格好良くないが、頭の１件を見てデータ型を判定する
            Select Case dt.Rows(1)(bd.fieldName).GetType().Name
                Case "DateTime"
                    sourceValue = Convert.ToDateTime(bd.sourceValue)
                    distinationvalue = Convert.ToDateTime(bd.distinationValue)
                Case Else
                    sourceValue = bd.sourceValue
                    distinationvalue = bd.distinationValue
            End Select

            For i As Integer = 0 To dt.Rows.Count - 1
                If dt.Rows(i)(bd.fieldName) = sourceValue Then
                    dr = dt.NewRow
                    For Each c In dt.Columns
                        If c.ToString = bd.fieldName Then
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

            Trn.Commit()

            'result OK
            result = "O K"
            msg = bd.sourceValue & " --> " & bd.distinationValue

        Catch ex As Exception
            'result ERR
            result = "ERR"
            errmsg = ex.GetType().ToString() & " " & ex.Message

            Try
                Trn.Rollback()
            Catch ex2 As Exception
                errmsg = ex2.GetType().ToString() & " " & ex2.Message
                MsgBox("Rollback Exception Type: " & ex2.GetType().ToString)
                ' This catch block will handle any errors that may have occured
                ' on the server that would cause the rollback to fail, such as
                ' a closed connection
            End Try

        Finally
            Main = "(" & result & ")" & "[" & bd.datatableName & " / " _
                & bd.fieldName & "] " & msg & " " & errmsg & " " & err2msg
            Con.Close()
            Con = Nothing
        End Try
        '---------------------------------------------------------------------------
    End Function

    Private ReadOnly Property connectionString(ByVal bd As test2.testmodule.BData) As System.String
        Get
            Return "Data Source=" & bd.serverName & ";" _
                & "Initial Catalog=" & bd.databaseName & ";" _
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

        If test2.testmodule.checkSave = 0 Then
            Exit Sub
        End If

        sw = New System.IO.StreamWriter(
            test2.testmodule.saveFile, False, System.Text.Encoding.GetEncoding("SHIFT_JIS"))

        If Not bd.serverName = vbNullString Then
            sw.WriteLine(bd.serverName)
        Else
            sw.WriteLine("サーバ名")
        End If

        If Not bd.databaseName = vbNullString Then
            sw.WriteLine(bd.databaseName)
        Else
            sw.WriteLine("データベース名")
        End If

        If Not bd.fieldName = vbNullString Then
            sw.WriteLine(bd.fieldName)
        Else
            sw.WriteLine("フィールド名")
        End If

        If Not bd.sourceValue = vbNullString Then
            sw.WriteLine(bd.sourceValue)
        Else
            sw.WriteLine("複製元データ")
        End If

        If Not bd.distinationValue = vbNullString Then
            sw.WriteLine(bd.distinationValue)
        Else
            sw.WriteLine("複製後データ")
        End If

        sw.Close()

        sw = Nothing
    End Sub

    Private Function checkSave() As System.Int32
        If Not System.IO.Directory.Exists(test2.testmodule.saveDir) Then
            System.IO.Directory.CreateDirectory(test2.testmodule.saveDir)
            System.IO.File.Create(test2.testmodule.iniFile)
        End If

        '既に同名フォルダーが存在する場合、作成しない
        If Not System.IO.File.Exists(test2.testmodule.iniFile) Then
            checkSave = 0
            Exit Function
        End If

        'saveFileが存在しない
        If Not System.IO.File.Exists(test2.testmodule.saveFile) Then
            System.IO.File.Create(test2.testmodule.saveFile)
            checkSave = 2
            Exit Function
        End If

        '正常
        checkSave = 1
    End Function

    Public Function Load(ByVal bd As test2.testmodule.BData)
        Dim txt(4) As System.String
        Dim i As System.Int32 : i = 0

        Load = bd
        If checkSave() = (0 Or 2) Then
            Exit Function
        End If

        sr = New System.IO.StreamReader(
            test2.testmodule.saveFile, System.Text.Encoding.GetEncoding("SHIFT_JIS"))

        While Not sr.EndOfStream
            If i >= 5 Then
                Continue While
            End If

            txt(i) = sr.ReadLine

            i += 1
        End While

        If i = UBound(txt) + 1 Then
            bd.serverName = txt(0)
            bd.databaseName = txt(1)
            bd.fieldName = txt(2)
            bd.sourceValue = txt(3)
            bd.distinationValue = txt(4)
            Load = bd
        End If

        sr.Close()

        sr = Nothing
    End Function

    Public Function changeFontSize(txtSize As System.Int32, winHeight As System.Int32) As System.Int32
        changeFontSize = (txtSize / test2.testmodule.defaultWinHeight) * winHeight
    End Function
End Module
