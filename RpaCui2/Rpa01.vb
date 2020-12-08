Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Runtime.InteropServices
Imports Outlook = Microsoft.Office.Interop.Outlook

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

    ' 依頼書定義
    Public Class Iraisho
        Public BookName As String
        Public MaxCount As Integer
        Public TargetPattern As List(Of String)
        Public ResheetType As String
        Public BankNameCell As String
        Public TantouCell As String
        Public ItakusyacodeCell As String
        Public ZieisinkincodeCell As String
        Public HikiotoshikinyukikanbangouCell As String
        Public MeisaiTableTopRow As Integer
        Public MeisaiTableBottomRow As Integer
        Public SitenmeiColumns As Integer
        Public SitencodeColumns As Integer
        Public KokyakucodeColumns As Integer
        Public FunocodeColumns As Integer      ' 足利専用プロパティ
        Public ShubetColumns As Integer
        Public KouzabangouColumns As Integer
        Public KouzameigiColumns As Integer
        Public SeikyugakuColumns As Integer
        Public TuchoukigouColumns As Integer     ' ゆうちょ専用プロパティ
        Public TuchoubangouColumns As Integer    ' ゆうちょ専用プロパティ
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

    ' 送付明細定義    
    Public Class SofuMeisai
        Public Kokyakucode As String
        Public Kouzameigi As String
        Public Furikaekingaku As String
        Public Ginkoucode As String
        Public Sitencode As String
        Public Ginkoumei As String
        Public MeisaiString As String
        Public HeaderFlag As Boolean
        Public ResheetFlag As Boolean
    End Class

    Private _SofuMeisaiDatas As List(Of SofuMeisai)
    <JsonIgnore>
    Public Property SofuMeisaiDatas As List(Of SofuMeisai)
        Get
            If Me._SofuMeisaiDatas Is Nothing Then
                Me._SofuMeisaiDatas = New List(Of SofuMeisai)
            End If
            Return Me._SofuMeisaiDatas
        End Get
        Set(value As List(Of SofuMeisai))
            Me._SofuMeisaiDatas = value
        End Set
    End Property

    Private __AttachmentFileName As String
    Private _AttachmentFileName As String
    Public Property AttachmentFileName As String
        Get
            Return Me._AttachmentFileName
        End Get
        Set(value As String)
            Me._AttachmentFileName = value
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
        Dim atcfile = vbNullString    '添付ファイル（フルパス）
        Dim bname = vbNullString      '添付ファイル（ファイル名（拡張子除く）のみ）
        Dim infile = vbNullString     '解凍後ファイル名
        Dim incsv = Rpa.MyProjectWorkDirectory & "\input.csv"

        Console.WriteLine("添付ファイルを検索しています...")

        ' 添付ファイルを保存し、その名前を取得
        Me.__AttachmentFileName = _GetAttachmentFile()
        For Each f In Directory.GetFiles(Rpa.MyProjectWorkDirectory)
            If Path.GetFileName(f) = Me.__AttachmentFileName Then
                atcfile = f
            End If
        Next
        If String.IsNullOrEmpty(atcfile) Then
            Console.WriteLine("エラー：添付ファイルが見つかりません")
            Return 1000
        End If
        bname = Path.GetFileNameWithoutExtension(atcfile)
        infile = Rpa.MyProjectWorkDirectory & "\" & bname & ".xls"

        ' 添付ファイルを解凍
        Console.WriteLine("添付ファイルを解凍します...")
        If Not File.Exists(Me.AttacheCase) Then
            Console.WriteLine($"エラー：アタッシュケース '{Me.AttacheCase}' がありません")
            Return 1000
        End If
        Call Rpa.RunShell(Me.AttacheCase, $"/c {atcfile} /p={Me.PasswordOfAttacheCase} /de=1 /ow=0 /opf=0 /exit=1")
        If Not File.Exists(infile) Then
            Console.WriteLine($"エラー：解凍後ファイル '{infile}' がありません")
            Return 1000
        End If

        ' ＣＳＶデータ生成
        Rpa.InvokeMacro("Rpa01.CreateInputTextData", {infile, incsv})
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

        Call _CreateSofuMeisaiDatas(tincsv, tincsv2)
        '-----------------------------------------------------------------------------------------'


        ' 各停止依頼書を作成
        '-----------------------------------------------------------------------------------------'
        'Dim fname As String
        'For Each f In Directory.GetFiles(Me.IraishoDirectory)
        '    fname = Path.GetFileName(f)
        '    File.Copy(f, $"{Me.Work2Directory}\{fname}", True)
        'Next
        '-----------------------------------------------------------------------------------------'


        ' 各停止依頼書ファイルをコピーする
        '-----------------------------------------------------------------------------------------'
        '-----------------------------------------------------------------------------------------'
    End Function


    Private Sub _CreateSofuMeisaiDatas(ByVal infile As String, ByVal otfile As String)
        Dim slist As New List(Of String)
        Dim sr As StreamReader
        Dim sw As StreamWriter
        Dim cnt As Integer
        Dim idx1 As Integer
        Dim work As String
        Dim v1 As String
        Dim v2 As String
        Dim vs1() As String
        Dim vs2() As String
        Dim bank As String
        Dim line As String
        Dim meisai As SofuMeisai
        Dim zsmd As SofuMeisai
        Dim miraisho As Iraisho
        Dim R_iraishos = CType(Rpa.RootProjectObject.IraishoDatas, List(Of Iraisho))
        sr = New StreamReader(infile, System.Text.Encoding.GetEncoding(MyEncoding))
        sw = New StreamWriter(otfile, False, System.Text.Encoding.GetEncoding(MyEncoding))

        ' ヘッダーは読み飛ばし
        sr.ReadLine()

        Do Until sr.EndOfStream
            line = sr.ReadLine()
            vs1 = line.Split(",")
            meisai = New SofuMeisai With {
                .MeisaiString = line,
                .Kokyakucode = vs1(4),
                .Ginkoucode = vs1(7),
                .Sitencode = vs1(8),
                .Ginkoumei = vs1(9),
                .Kouzameigi = vs1(13),
                .Furikaekingaku = vs1(14)
            }
            Me.SofuMeisaiDatas.Add(meisai)
        Loop

        ' 銀行コード－＞支店コードでソート
        Me.SofuMeisaiDatas.Sort(
            Function(a, b)
                If a.Ginkoucode <> b.Ginkoucode Then
                    Return (a.Ginkoucode - b.Ginkoucode)
                Else
                    Return (a.Sitencode - b.Sitencode)
                End If
            End Function
        )

        ' 依頼書の対象銀行コードが合致する場合、その依頼書データを取得
        Dim smdcount = 0
        zsmd = New SofuMeisai
        For Each smd In Me.SofuMeisaiDatas
            If zsmd.Ginkoucode <> smd.Ginkoucode Then
                smdcount = 0
                miraisho = Nothing
                miraisho = R_iraishos.Find(
                    Function(ri)
                        Dim hit = ri.TargetBanks.Find(
                            Function(tb)
                                Return (tb.Code = smd.Ginkoucode)
                            End Function
                        )
                        Return IIf(hit IsNot Nothing, True, False)
                    End Function
                )
                If miraisho Is Nothing Then
                    miraisho = R_iraishos.Find(
                        Function(ri)
                            If ri.TargetPattern IsNot Nothing Then
                                Dim hit = ri.TargetPattern.Find(
                                    Function(tp)
                                        Return smd.Ginkoumei.Contains(tp)
                                    End Function
                                )
                                Return IIf(hit IsNot Nothing, True, False)
                            Else
                                Return False
                            End If
                        End Function
                    )
                End If
                If miraisho Is Nothing Then
                    miraisho = R_iraishos(0)
                End If
                smd.HeaderFlag = True
            End If

            smdcount += 1
            smd.ResheetFlag = IIf(smdcount > miraisho.MaxCount, True, False)
            smdcount = IIf(smd.ResheetFlag, 1, smdcount)

            ' 各種依頼書情報付加
            smd.MeisaiString &= IIf(smd.HeaderFlag, "1", "0")                        ' idx21
            smd.MeisaiString &= "," & IIf(smd.ResheetFlag, "1", "0")                 ' idx22
            smd.MeisaiString &= "," & miraisho.BookName                              ' idx23
            smd.MeisaiString &= "," & miraisho.ResheetType                           ' idx24
            smd.MeisaiString &= "," & miraisho.BankNameCell                          ' idx25
            smd.MeisaiString &= "," & miraisho.TantouCell                            ' idx26
            smd.MeisaiString &= "," & miraisho.ItakusyacodeCell                      ' idx27
            smd.MeisaiString &= "," & miraisho.ZieisinkincodeCell                    ' idx28
            smd.MeisaiString &= "," & miraisho.HikiotoshikinyukikanbangouCell        ' idx29
            smd.MeisaiString &= "," & miraisho.MeisaiTableTopRow                     ' idx30
            smd.MeisaiString &= "," & miraisho.MeisaiTableBottomRow                  ' idx31
            smd.MeisaiString &= "," & (miraisho.MeisaiTableTopRow + smdcount - 1)    ' idx32
            smd.MeisaiString &= "," & miraisho.SitenmeiColumns                       ' idx33
            smd.MeisaiString &= "," & miraisho.SitencodeColumns                      ' idx34
            smd.MeisaiString &= "," & miraisho.KokyakucodeColumns                    ' idx35
            smd.MeisaiString &= "," & miraisho.FunocodeColumns                       ' idx36
            smd.MeisaiString &= "," & miraisho.ShubetColumns                         ' idx37
            smd.MeisaiString &= "," & miraisho.KouzabangouColumns                    ' idx38
            smd.MeisaiString &= "," & miraisho.KouzameigiColumns                     ' idx39
            smd.MeisaiString &= "," & miraisho.SeikyugakuColumns                     ' idx30
            smd.MeisaiString &= "," & miraisho.TuchoukigouColumns                    ' idx41
            smd.MeisaiString &= "," & miraisho.TuchoubangouColumns                   ' idx42
            smd.MeisaiString &= "," & miraisho.BikouColumns                          ' idx43

            sw.WriteLine(smd.MeisaiString)
            zsmd = smd
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


    'Private Sub _CreateSofuMeisaiDatas(ByVal infile As String, ByVal otfile As String)
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


    'Private Sub _RunCmd(ByVal arg As String)
    '    Dim proc = New System.Diagnostics.Process()
    '    proc.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
    '    proc.StartInfo.UseShellExecute = False
    '    proc.StartInfo.RedirectStandardOutput = True
    '    proc.StartInfo.RedirectStandardInput = False
    '    proc.StartInfo.CreateNoWindow = True
    '    proc.StartInfo.Arguments = arg
    '    proc.Start()
    '    proc.WaitForExit()
    '    proc.Close()
    'End Sub

    Private Function _GetAttachmentFile() As String
        Const MAPI = "MAPI"
        Dim olapp, ns, inbox, sfolder, bfolder, [item]
        Dim fnc As Func(Of Object, String)
        Dim rtn = vbNullString
        Try
            olapp = CreateObject("Outlook.Application")
            ns = olapp.GetNamespace(MAPI)
            inbox = ns.GetDefaultFolder(6)
            sfolder = inbox.Folders.Item(Me.OutlookSourceInboxName)
            bfolder = inbox.Folders.Item(Me.OutlookBackupInboxName)
            For Each i In sfolder.Items
                If [item] Is Nothing Then
                    [item] = i
                Else
                    If [item].CreationTime < i.CreationTime Then
                        [item] = i
                    End If
                End If
            Next

            ' 添付ファイルを保存し、その名前を得る
            fnc = Function(atc As Object)
                      Call atc.SaveAsFile($"{Rpa.MyProjectWorkDirectory}\{atc.FileName}")
                      Return atc.FileName
                  End Function
            If [item] IsNot Nothing Then
                For Each attachment In item.Attachments
                    If String.IsNullOrEmpty(Me.AttachmentFileName) Then
                        rtn = fnc(attachment)
                    Else
                        If attachment.FileName = Me.AttachmentFileName Then
                            rtn = fnc(attachment)
                        End If
                    End If
                Next
            End If

            Do Until sfolder.Items.Count = 0
                For Each i In sfolder.Items
                    Call i.Move(bfolder)
                    Exit For
                Next
            Loop
        Finally
            Marshal.ReleaseComObject(olapp)
        End Try
        Return rtn
    End Function
End Class
