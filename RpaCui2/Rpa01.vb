Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Rpa01 : Inherits RpaBase(Of Rpa01)
    Private Const EXCEL_APP As String = "Excel.Application"
    Private Const CSCRIPT As String = "cscript"
    Private Const SHIFT_JIS = "Shift-JIS"

    Private _ItakusyaCodeDictionary As Dictionary(Of String, String)
    Public Property ItakusyaCodeDictionary As Dictionary(Of String, String)
        Get
            If Me._ItakusyaCodeDictionary Is Nothing Then
                Me._ItakusyaCodeDictionary = New Dictionary(Of String, String)
                Me._ItakusyaCodeDictionary.Add("Moteki", "0029005")
            End If
            Return Me._ItakusyaCodeDictionary
        End Get
        Set(value As Dictionary(Of String, String))
            Me._ItakusyaCodeDictionary = value
        End Set
    End Property

    Private _IraishoDirectory As String
    <JsonIgnore>
    Public ReadOnly Property IraishoDirectory As String
        Get
            If String.IsNullOrEmpty(Me._IraishoDirectory) Then
                Me._IraishoDirectory = Rpa.MyProjectDirectory & "\iraisho"
                If Not Directory.Exists(Me._IraishoDirectory) Then
                    Directory.CreateDirectory(Me._IraishoDirectory)
                End If
            End If
            Return Me._IraishoDirectory
        End Get
    End Property

    Private _Work2Directory As String
    <JsonIgnore>
    Public ReadOnly Property Work2Directory As String
        Get
            If String.IsNullOrEmpty(Me._Work2Directory) Then
                Me._Work2Directory = Rpa.MyProjectDirectory & "\work2"
                If Not Directory.Exists(Me._Work2Directory) Then
                    Directory.CreateDirectory(Me._Work2Directory)
                End If
            End If
            Return Me._Work2Directory
        End Get
    End Property

    Public Class Iraisho
        Public BookName As String
        Public LayoutType As Integer
        Public MaxCount As Integer
        Public ResheetType As Integer
        Public BankNameCell As String
        Public TantouCell As String
        Public ItakusyacodeCell As String
        Public ZieisinkincodeCell As String
        Public HikiotoshikinyukikanbangoCell As String
        Public SitenmeiColumns As Integer
        Public SitencodeColumns As Integer
        Public KokyakucodeColumns As Integer
        Public FunocodeColumns As Integer      ' 足利専用プロパティ
        Public ShubetColumns As Integer
        Public KouzabangoColumns As Integer
        Public KouzameigiColumns As Integer
        Public SeikyugakColumns As Integer
        Public TuchokigoColumns As Integer     ' ゆうちょ専用プロパティ
        Public BikouColumns As Integer
        Private _TargetBanks As List(Of BankInfo)
        Public Property TargetBanks As List(Of BankInfo)
            Get
                If Me._TargetBanks Is Nothing Then
                    Me._TargetBanks = New List(Of BankInfo)
                End If
                Return Me._TargetBanks
            End Get
            Set(value As List(Of BankInfo))
                Me._TargetBanks = value
            End Set
        End Property
        Public Class BankInfo
            Public Code As String
            Public Name As String
            Public ZieiSinkinCode As String
        End Class
    End Class

    Private _IraishoDatas As List(Of Iraisho)
    Public Property IraishoDatas As List(Of Iraisho)
        Get
            If Me._IraishoDatas Is Nothing Then
                Me._IraishoDatas = New List(Of Iraisho)
            End If
            Return Me._IraishoDatas
        End Get
        Set(value As List(Of Iraisho))
            Me._IraishoDatas = value
        End Set
    End Property

    Private _PasswordOfAttacheCase As String
    Public Property PasswordOfAttacheCase As String
        Get
            Return Me._PasswordOfAttacheCase
        End Get
        Set(value As String)
            Me._PasswordOfAttacheCase = value
        End Set
    End Property

    Private _OutlookSourceInboxName As String
    Public Property OutlookSourceInboxName As String
        Get
            Return Me._OutlookSourceInboxName
        End Get
        Set(value As String)
            Me._OutlookSourceInboxName = value
        End Set
    End Property

    Private _OutlookBackupInboxName As String
    Public Property OutlookBackupInboxName As String
        Get
            Return Me._OutlookBackupInboxName
        End Get
        Set(value As String)
            Me._OutlookBackupInboxName = value
        End Set
    End Property

    Private _MasterCsvFileName As String
    Public Property MasterCsvFileName As String
        Get
            If String.IsNullOrEmpty(Me._MasterCsvFileName) Then
                Me._MasterCsvFileName = "master.csv"
            End If
            Return Me._MasterCsvFileName
        End Get
        Set(value As String)
            Me._MasterCsvFileName = value
        End Set
    End Property

    Private _AttacheCase As String
    Public Property AttacheCase As String
        Get
            Return Me._AttacheCase
        End Get
        Set(value As String)
            Me._AttacheCase = value
        End Set
    End Property

    Private _RestartCode As String
    Public Property RestartCode As String
        Get
            Return Me._RestartCode
        End Get
        Set(value As String)
            Me._RestartCode = value
        End Set
    End Property

    Public Overrides Function SetupProjectObject(project As String) As Object
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Main() As Integer
        Call Rpa.ModelSave(Rpa.SystemJsonFileName, Rpa)

        ' マスターファイルのチェック・コピー
        '-----------------------------------------------------------------------------------------'
        Dim inmaster = Rpa.MyProjectDirectory & "\" & Me.MasterCsvFileName
        Dim wkmaster = Rpa.MyProjectWorkDirectory & "\" & Me.MasterCsvFileName
        Console.WriteLine("以下のファイルを用意したら、[Enter]キーをクリックしてください")
        Console.WriteLine("ファイル名 : " & Me.MasterCsvFileName)
        Console.ReadLine()
        If Not File.Exists(inmaster) Then
            Console.WriteLine($"エラー：ファイル '{inmaster}' が存在しません")
            Return 1000
        End If
        File.Copy(inmaster, wkmaster, True)
        '-----------------------------------------------------------------------------------------'



        ' 添付ファイルの取得・解凍・ＣＳＶデータ生成
        '-----------------------------------------------------------------------------------------'
        'Dim atcfile = vbNullString                                        '添付ファイル
        'Dim bname = Path.GetFileNameWithoutExtension(atcfile)             '添付ファイルのファイル名(拡張子なし)
        'Dim infile = Rpa.MyProjectWorkDirectory & "\" & bname & ".xls"    '解凍後ファイル
        Dim incsv = Rpa.MyProjectWorkDirectory & "\input.csv"

        'Console.WriteLine("添付ファイルを検索しています...")
        'Call _GetAttachmentFile()
        'For Each f In Directory.GetFiles(Rpa.MyProjectWorkDirectory)
        '    If Path.GetExtension(f) = "atc" Then
        '        atcfile = f
        '    End If
        'Next
        'If String.IsNullOrEmpty(atcfile) Then
        '    Console.WriteLine("エラー：添付ファイルが見つかりません")
        '    Return 1000
        'End If
        'Console.WriteLine("添付ファイルを解凍します...")
        'If File.Exists(Me.AttacheCase) Then
        '    Console.WriteLine($"エラー：アタッシュケース '{Me.AttacheCase}' がありません")
        '    Return 1000
        'End If
        'Call _RunCmd("/c " _
        '           & "/p=${Me.AttacheCase} " _
        '           & "/de=1 /ow=0 /opf=0 /exit=1")

        'If Not File.Exists(infile) Then
        '    Console.WriteLine("エラー：解凍後ファイル '${infile}' がありません")
        '    Return 1000
        'End If

        'Dim args() = {infile, incsv}
        'Rpa.InvokeMacro("Rpa01.CreateInputTextData", args)
        '-----------------------------------------------------------------------------------------'


        ' ＣＳＶから対象得意先（モテキ）のみ抜き出し・入力ＣＳＶとのマッチング
        '-----------------------------------------------------------------------------------------'
        Dim inmaster2 = inmaster
        Dim tmpcsv = Rpa.MyProjectWorkDirectory & "\tmp.csv"
        Dim target = Me.ItakusyaCodeDictionary("Moteki")
        Dim incsv2 = incsv
        Dim tincsv = Rpa.MyProjectWorkDirectory & "\t_input.csv"
        Dim tincsv2 = Rpa.MyProjectWorkDirectory & "\t_input2.csv"
        Call _CreateTmpCsv(inmaster2, target, tmpcsv)
        Call _CompareInputCsvToTmpCsv(incsv, tmpcsv, tincsv)

        Call _AddTeishiIraishoMetaDatas(tincsv, tincsv2)
        '-----------------------------------------------------------------------------------------'


        ' 各停止依頼書を作成
        '-----------------------------------------------------------------------------------------'
        Dim fname As String
        For Each f In Directory.GetFiles(Me.IraishoDirectory)
            fname = Path.GetFileName(f)
            File.Copy(f, $"{Me.Work2Directory}\{fname}", True)
        Next
        '-----------------------------------------------------------------------------------------'


        ' 各停止依頼書ファイルをコピーする
        '-----------------------------------------------------------------------------------------'
        '-----------------------------------------------------------------------------------------'
    End Function


    Private Sub _AddTeishiIraishoMetaDatas(ByVal infile As String, ByVal otfile As String)
        Dim slist As New List(Of String)
        Dim sr As StreamReader
        Dim sw As StreamWriter
        Dim cnt As Integer
        Dim idx1 As Integer
        Dim work As String
        Dim v1 As String
        Dim v2 As String
        Dim bank As String
        sr = New StreamReader(infile, System.Text.Encoding.GetEncoding(MyEncoding))
        sw = New StreamWriter(otfile, False, System.Text.Encoding.GetEncoding(MyEncoding))

        ' ヘッダー
        sr.ReadLine()

        Do Until sr.EndOfStream
            slist.Add(sr.ReadLine())
        Loop

        ' 簡易的に銀行コードでバブルソート
        ' コードはあまりきれいじゃないので時間があるときに見直したい・・・
        cnt = slist.IndexOf(slist.Last)
        For idx1 = 0 To (cnt - 1)
            For idx2 = 0 To (cnt - (idx1 + 1))
                v1 = (slist(idx2).Split(","))(7)
                v2 = (slist(idx2 + 1).Split(","))(7)
                If v1 > v2 Then
                    work = slist(idx2)
                    slist(idx2) = slist(idx2 + 1)
                    slist(idx2 + 1) = work
                End If
            Next
        Next

        For Each s In slist
            sw.WriteLine(s)
        Next

        If sr IsNot Nothing Then
            sr.Close()
            sr.Dispose()
        End If
        If sw IsNot Nothing Then
            sw.Close()
            sw.Dispose()
        End If
    End Sub


    'Private Sub _AddTeishiIraishoMetaDatas(ByVal infile As String, ByVal otfile As String)
    '    Dim sr As StreamReader
    '    Dim sw As StreamWriter
    '    Dim line As String
    '    Dim v() As String
    '    sr = New StreamReader(infile, System.Text.Encoding.GetEncoding(MyEncoding))
    '    sw = New StreamWriter(otfile, False, System.Text.Encoding.GetEncoding(MyEncoding))

    '    ' ヘッダー
    '    sw.WriteLine(line)

    '    Do Until sr.EndOfStream
    '        line = sr.ReadLine
    '        v = line.Split(",")
    '        If v(2) = target Then
    '            sw.WriteLine(line)
    '        End If
    '    Loop

    '    If sr IsNot Nothing Then
    '        sr.Close()
    '        sr.Dispose()
    '    End If
    '    If sw IsNot Nothing Then
    '        sw.Close()
    '        sw.Dispose()
    '    End If
    'End Sub


    Private Sub _CompareInputCsvToTmpCsv(ByVal infile As String, ByVal tmpcsv As String, ByVal otfile As String)
        Dim isr As StreamReader
        Dim tsr As StreamReader
        Dim sw As StreamWriter
        Dim iline As String
        Dim tline As String
        Dim iv() As String
        Dim tv() As String
        Dim errlevel = 0
        Dim oline As String
        Dim break = False
        Dim restart = IIf(String.IsNullOrEmpty(Me.RestartCode), False, True)
        Dim errmsg(15) As String
        errmsg(4) = "顧客コードなし"
        errmsg(13) = "口座名義不一致"
        errmsg(14) = "振替金額不一致"
        errmsg(15) = "チェックＯＫ"

        isr = New StreamReader(infile, System.Text.Encoding.GetEncoding(MyEncoding))
        tsr = New StreamReader(tmpcsv, System.Text.Encoding.GetEncoding(MyEncoding))
        sw = New StreamWriter(otfile, False, System.Text.Encoding.GetEncoding(MyEncoding))

        ' ヘッダー
        tline = tsr.ReadLine
        tline = tline.Replace("""", vbNullString)
        tline = tline.Replace("=", vbNullString)
        sw.WriteLine(tline)

        Do Until isr.EndOfStream
            iline = isr.ReadLine
            iline = iline.Replace(" ", vbNullString)
            iv = iline.Split(",")
            iv(14) = iv(14).Replace("\", vbNullString)

            If restart Then
                If Me.RestartCode = iv(4) Then
                    restart = False
                    Continue Do
                Else
                    Continue Do
                End If
            End If

            break = False
            errlevel = 4
            Do Until tsr.EndOfStream
                tline = tsr.ReadLine
                tline = tline.Replace("""", vbNullString)
                tline = tline.Replace("=", vbNullString)
                tv = tline.Split(",")
                tv(4) = tv(4).Trim
                tv(4) = tv(4).TrimStart("0")
                tv(4) = tv(4).PadLeft(8, "0")
                tv(13) = tv(13).Trim
                tv(13) = tv(13).Replace(" ", vbNullString)
                tv(14) = tv(14).Trim
                tv(14) = tv(14).Replace(" ", vbNullString)

                If iv(4) = tv(4) Then
                    break = True
                    errlevel = 13
                    If iv(13) = tv(13) Then
                        errlevel = 14
                        If iv(14) = tv(14) Then
                            errlevel = 15
                        End If
                    End If
                End If

                If break Then
                    Exit Do
                End If
            Loop

            oline = vbNullString
            If break Then
                For Each v In tv
                    v = v.TrimStart
                    v = v.TrimEnd
                    oline &= v & ","
                Next
                oline &= errlevel.ToString & "," & errmsg(errlevel) & ","
                sw.WriteLine(oline)
            Else
                Console.WriteLine("以下のデータは送付明細上に存在しません")
                Console.WriteLine($"データ: {iline.TrimEnd}")
            End If

            tsr.BaseStream.Seek(0, SeekOrigin.Begin)
            tsr.ReadLine()
        Loop

        If isr IsNot Nothing Then
            isr.Close()
            isr.Dispose()
        End If
        If tsr IsNot Nothing Then
            tsr.Close()
            tsr.Dispose()
        End If
        If sw IsNot Nothing Then
            sw.Close()
            sw.Dispose()
        End If
    End Sub


    Private Sub _CreateTmpCsv(ByVal infile As String, ByVal target As String, ByVal otfile As String)
        Dim sr As StreamReader
        Dim sw As StreamWriter
        Dim line As String
        Dim v() As String
        sr = New StreamReader(infile, System.Text.Encoding.GetEncoding(SHIFT_JIS))
        sw = New StreamWriter(otfile, False, System.Text.Encoding.GetEncoding(MyEncoding))

        ' ヘッダー
        line = sr.ReadLine
        line = line.Replace("""", vbNullString)
        line = line.Replace("=", vbNullString)
        sw.WriteLine(line)

        Do Until sr.EndOfStream
            line = sr.ReadLine
            v = line.Split(",")
            If v(2) = target Then
                sw.WriteLine(line)
            End If
        Loop

        If sr IsNot Nothing Then
            sr.Close()
            sr.Dispose()
        End If
        If sw IsNot Nothing Then
            sw.Close()
            sw.Dispose()
        End If
    End Sub


    Private Sub _RunCmd(ByVal arg As String)
        Dim proc = New System.Diagnostics.Process()
        proc.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
        proc.StartInfo.UseShellExecute = False
        proc.StartInfo.RedirectStandardOutput = True
        proc.StartInfo.RedirectStandardInput = False
        proc.StartInfo.CreateNoWindow = True
        proc.StartInfo.Arguments = arg
        proc.Start()
        proc.WaitForExit()
        proc.Close()
    End Sub

    Private Sub _GetAttachmentFile()
        Const MAPI = "MAPI"
        Const HIZUKE = "日付"
        Dim ol = CreateObject("Outlook.Application")
        Dim ns = ol.GetNamespace(MAPI)
        Dim inbox = ns.GetDefaultFolder(6)
        Dim sfolder = inbox.Folders.Item(Me.OutlookSourceInboxName)
        Dim bfolder = inbox.Folders.Item(Me.OutlookBackupInboxName)
        Dim afile As String

        sfolder.Items.Sort(HIZUKE, True)
        For Each [item] In sfolder.Items
            If [item].Attachments.Count = 0 Then
                Continue For
            End If
            For Each attachment In [item].Attachments
                afile = Rpa.MyProjectWorkDirectory & "\" & attachment
                attachment.SaveAsFile(afile)
            Next
        Next

        Do Until sfolder.Items.Count = 0
            sfolder.Items(0).Move(bfolder)
        Loop
    End Sub
End Class
