﻿Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Runtime.InteropServices
Imports System.Globalization

Public Class Rpa01 : Inherits RpaBase(Of Rpa01)
    Private Const EXCEL_APP As String = "Excel.Application"
    Private Const CSCRIPT As String = "cscript"
    Private Const SHIFT_JIS = "Shift-JIS"

    Public Syunoukigyoucode As String              ' 収納企業コード
    Public Syunoukigyoumei As String               ' 収納企業名
    Public Syunoukigyouaddress As String           ' 収納企業住所
    Public Syunoukigyoutel As String               ' 収納企業電話番号
    Public Syunoukigyoutantousya As String         ' 収納企業担当者
    Public Haraikomisakikouzabangou As String      ' 払込先口座番号
    Public Haraikomikinshubetu As String           ' 払込金種別

    ' 加工済送付明細のカラムサイズ
    Public SofuMeisaiColumnALength As Double
    Public SofuMeisaiColumnBLength As Double
    Public SofuMeisaiColumnCLength As Double
    Public SofuMeisaiColumnDLength As Double
    Public SofuMeisaiColumnELength As Double
    Public SofuMeisaiColumnFLength As Double
    Public SofuMeisaiColumnGLength As Double
    Public SofuMeisaiColumnHLength As Double
    Public SofuMeisaiColumnILength As Double
    Public SofuMeisaiColumnJLength As Double
    Public SofuMeisaiColumnKLength As Double
    Public SofuMeisaiColumnLLength As Double
    Public SofuMeisaiColumnMLength As Double
    Public SofuMeisaiColumnNLength As Double
    Public SofuMeisaiColumnOLength As Double
    Public SofuMeisaiColumnPLength As Double
    Public SofuMeisaiColumnQLength As Double
    Public SofuMeisaiColumnRLength As Double
    Public SofuMeisaiColumnSLength As Double
    Public SofuMeisaiColumnTLength As Double
    Public SofuMeisaiColumnULength As Double

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
        Public SourceSheetName As String

        ' 停止依頼元
        Public SyunoukigyoucodeCell As String             ' 収納企業コード
        Public SyunoukigyoumeiCell As String              ' 収納企業名
        Public SyunoukigyouaddressCell As String          ' 収納企業住所
        Public SyunoukigyoutelCell As String              ' 収納企業電話番号
        Public SyunoukigyoutantousyaCell As String        ' 収納企業担当者
        Public HaraikomisakikouzabangouCell As String     ' 払込先口座番号
        Public HaraikomikinshubetuCell As String          ' 払込金種別
        Public IraibiCell As String                       ' 依頼日
        Public SakuseibiCell As String                    ' 作成日（８桁）
        Public SakuseibiYYYYCell As String                ' 作成年（４桁）
        Public SakuseibiYYCell As String                  ' 作成年（下２桁）
        Public SakuseibiWaYYCell As String                ' 作成年（和暦）
        Public SakuseibiWaZYCell As String                ' 作成年（和暦、ゼロサプレス）
        Public SakuseibiMMCell As String                  ' 作成月（２桁）
        Public SakuseibiZMCell As String                  ' 作成月（ゼロサプレス）
        Public SakuseibiDDCell As String                  ' 作成日（２桁）
        Public SakuseibiZDCell As String                  ' 作成日（ゼロサプレス）
        Public FurikaesiteibiCell As String               ' 振替指定日

        ' 停止依頼元
        Public ItakusyacodeCell As String                 ' 委託者コード
        Public ItakusyameiCell As String                  ' 委託者名

        ' 停止依頼先
        Public Iraisaki As String                         ' 提出先
        Public IraisakiCell As String                     ' 提出先
        Public I_Ginkoucode As String                     ' 銀行コード
        Public I_GinkoucodeCell As String                 ' 銀行コード
        Public I_Ginkoucode1 As String                    ' 銀行コード（左から１桁目）
        Public I_Ginkoucode1Cell As String                ' 銀行コード（左から１桁目）
        Public I_Ginkoucode2 As String                    ' 銀行コード（左から２桁目）
        Public I_Ginkoucode2Cell As String                ' 銀行コード（左から２桁目）
        Public I_Ginkoucode3 As String                    ' 銀行コード（左から３桁目）
        Public I_Ginkoucode3Cell As String                ' 銀行コード（左から３桁目）
        Public I_Ginkoucode4 As String                    ' 銀行コード（左から４桁目）
        Public I_Ginkoucode4Cell As String                ' 銀行コード（左から４桁目）
        Public I_Ginkoumei As String                      ' 銀行名
        Public I_GinkoumeiCell As String                  ' 銀行名
        Public I_Ginkou As String                         ' 銀行名（銀行名から「銀行」などを除いた名称）
        Public I_GinkouCell As String                     ' 銀行名（銀行名から「銀行」などを除いた名称）
        Public I_GinkouSaffix As String                   ' 「銀行」・「信用金庫」など
        Public I_GinkouSaffixCell As String               ' 「銀行」・「信用金庫」など
        Public I_Zieisinkincode As String                 ' 自営信金コード
        Public I_ZieisinkincodeCell As String             ' 自営信金コード
        Public Baitaimei As String                        ' 媒体名
        Public BaitaimeiCell As String                    ' 媒体名

        ' 停止銀行
        Public T_GinkoucodeCell As String                 ' 銀行コード
        Public T_Ginkoucode1Cell As String                ' 銀行コード（左から１桁目）
        Public T_Ginkoucode2Cell As String                ' 銀行コード（左から２桁目）
        Public T_Ginkoucode3Cell As String                ' 銀行コード（左から３桁目）
        Public T_Ginkoucode4Cell As String                ' 銀行コード（左から４桁目）
        Public T_GinkoumeiCell As String                  ' 銀行名
        Public T_GinkouCell As String                     ' 銀行名（銀行名から「銀行」などを除いた名称）
        Public T_GinkouSaffixCell As String               ' 「銀行」・「信用金庫」など
        Public T_ZieisinkincodeCell As String             ' 自営信金コード

        ' 停止対象
        Public MeisaiTableTopRow As Integer

        Public KokyakucodeColumn As String                ' 顧客コード
        Public K_GinkoucodeColumn As String               ' 銀行コード
        Public SitencodeColumn As String                  ' 支店コード
        Public K_GinkoumeiColumn As String                ' 銀行名
        Public K_GinkouColumn As String                   ' 銀行名（銀行名から「銀行」などを除いた名称）
        Public SitenmeiColumn As String                   ' 支店名
        Public SitenColumn As String                      ' 支店名（支店名から「支店」を除いた名称）
        Public YokinshubetuColumn As String               ' 預金種別
        Public KouzabangouColumn As String                ' 口座番号
        Public TuchoukigouColumn As String                ' 通帳記号
        Public TuchoubangouColumn As String               ' 通帳番号
        Public KouzameigiColumn As String                 ' 口座名義（半角）
        Public JISKouzameigiColumn As String              ' 口座名義（全角）
        Public FurikaekingakuColumn As String             ' 振替金額

        Public Bikou As String                            ' 備考データ
        Public BikouColumn As String                      ' 備考
        Public Tekiyou As String                          ' 摘要データ
        Public TekiyouColumn As String                    ' 摘要
        Public Funoucode As String                        ' 不能コードデータ
        Public FunoucodeColumn As String                  ' 不能コード

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
            Public SourceSheetName As String
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

    ' 送付明細定義（ソートに必要な情報を定義）
    Public Class IraishoMeisai
        Public T_Ginkoucode As String                 ' 銀行コード
        Public T_Ginkoumei As String                  ' 銀行名
        Public Sitencode As String                    ' 支店コード
        Public MeisaiString As String
    End Class

    Private _IraishoMeisaiDatas As List(Of IraishoMeisai)
    <JsonIgnore>
    Public Property IraishoMeisaiDatas As List(Of IraishoMeisai)
        Get
            If Me._IraishoMeisaiDatas Is Nothing Then
                Me._IraishoMeisaiDatas = New List(Of IraishoMeisai)
            End If
            Return Me._IraishoMeisaiDatas
        End Get
        Set(value As List(Of IraishoMeisai))
            Me._IraishoMeisaiDatas = value
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

    Private _RestartCount As Integer
    Public Property RestartCount As Integer
        Get
            Return Me._RestartCount
        End Get
        Set(value As Integer)
            Me._RestartCount = value
        End Set
    End Property

    Public Overrides Function SetupProjectObject(project As String) As Object
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Main() As Integer
        Me.RestartCount = IIf(Me.RestartCount = 0, 1, Me.RestartCount)
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

        'Console.WriteLine("添付ファイルを検索しています...")

        '' 添付ファイルを保存し、その名前を取得
        'Me.__AttachmentFileName = _GetAttachmentFile()
        'For Each f In Directory.GetFiles(Rpa.MyProjectWorkDirectory)
        '    If Path.GetFileName(f) = Me.__AttachmentFileName Then
        '        atcfile = f
        '    End If
        'Next
        'If String.IsNullOrEmpty(atcfile) Then
        '    Console.WriteLine("エラー：添付ファイルが見つかりません")
        '    Return 1000
        'End If
        'bname = Path.GetFileNameWithoutExtension(atcfile)
        'infile = Rpa.MyProjectWorkDirectory & "\" & bname & ".xls"

        '' 添付ファイルを解凍
        'Console.WriteLine("添付ファイルを解凍します...")
        'If Not File.Exists(Me.AttacheCase) Then
        '    Console.WriteLine($"エラー：アタッシュケース '{Me.AttacheCase}' がありません")
        '    Return 1000
        'End If
        'Call Rpa.RunShell(Me.AttacheCase, $"/c {atcfile} /p={Me.PasswordOfAttacheCase} /de=1 /ow=0 /opf=0 /exit=1")
        'If Not File.Exists(infile) Then
        '    Console.WriteLine($"エラー：解凍後ファイル '{infile}' がありません")
        '    Return 1000
        'End If

        ' TEST TEST
        '-------------------------------------------------'
        infile = $"{Rpa.MyProjectDirectory}\input.xls"
        '-------------------------------------------------'

        ' ＣＳＶデータ生成
        Call Rpa.InvokeMacro("Rpa01.CreateInputTextData", {infile, incsv})
        '-----------------------------------------------------------------------------------------'


        ' ＣＳＶから対象得意先（モテキ）のみ抜き出し・入力ＣＳＶとのマッチング
        '-----------------------------------------------------------------------------------------'
        Dim target = Me.ItakusyaCodeDictionary("Moteki")
        Dim inmaster2 = inmaster
        Dim tmpcsv = $"{Rpa.MyProjectWorkDirectory}\tmp.csv"
        Dim incsv2 = incsv
        Dim tincsv = $"{Rpa.MyProjectWorkDirectory}\t_input.csv"
        Dim tincsv2 = $"{Rpa.MyProjectWorkDirectory}\t_input2.csv"
        Call _CreateTmpCsv(inmaster2, target, tmpcsv)
        Call _CompareInputCsvToTmpCsv(incsv, tmpcsv, tincsv)
        Call _CreateIraishoMeisaiDatas(tincsv, tincsv2)
        '-----------------------------------------------------------------------------------------'

        ' 各停止依頼書を作成
        '-----------------------------------------------------------------------------------------'
        Console.WriteLine("停止依頼書を作成中・・・")
        Dim tincsv3 = tincsv2
        Dim fname As String
        For Each f In Directory.GetFiles(Me.IraishoDirectory)
            fname = Path.GetFileName(f)
            File.Copy(f, $"{Me.Work2Directory}\{fname}", True)
        Next

        Call Rpa.InvokeMacro("Rpa01.CreateIraisho", {tincsv3})
        Console.WriteLine("停止依頼書作成完了！")
        '-----------------------------------------------------------------------------------------'

        ' 加工済送付明細の作成
        '-----------------------------------------------------------------------------------------'
        Console.WriteLine("加工済み送付明細を作成中・・・")
        Dim wkmaster_v2 = $"{Rpa.MyProjectWorkDirectory}\{Me.MasterCsvFileName}"
        Dim tincsv_v2 = $"{Rpa.MyProjectWorkDirectory}\t_input.csv"
        Dim outxlsx_1 = $"{Rpa.MyProjectBackupDirectory}\加工済送付明細.xlsx"
        Dim setting_v1 As String
        setting_v1 = Rpa.RootProjectObject.SofuMeisaiColumnALength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnBLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnCLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnDLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnELength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnFLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnGLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnHLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnILength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnJLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnKLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnLLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnMLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnNLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnOLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnPLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnQLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnRLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnSLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnTLength.ToString()
        setting_v1 &= "," & Rpa.RootProjectObject.SofuMeisaiColumnULength.ToString()
        If Me.RestartCount = 1 Then
            File.Delete(outxlsx_1)
            Call Rpa.InvokeMacro("Rpa01.CreateSofuMeisai", {wkmaster_v2, outxlsx_1, "master", "Shift-JIS"})
        End If
        Call Rpa.InvokeMacro("Rpa01.CreateSofuMeisai", {tincsv_v2, outxlsx_1, $"停止{Me.RestartCount.ToString}回目", "utf-8", setting_v1})
        Console.WriteLine("加工済送付明細作成完了！")
        '-----------------------------------------------------------------------------------------'


        '-----------------------------------------------------------------------------------------'
        Dim outtxt_1 = $"{Rpa.MyProjectBackupDirectory}\件数集計.txt"
        Console.WriteLine("件数集計ファイルを作成中・・・")
        Call _CreateSummaryFile(outtxt_1)
        Console.WriteLine("件数集計ファイル作成完了！")
        '-----------------------------------------------------------------------------------------'

        If Transaction.Parameters.Count > 0 Then
            If Transaction.Parameters.Last = "end" Then
                Me.RestartCode = vbNullString
                Me.RestartCount = 1
            End If
        Else
            Me.RestartCount += 1
        End If

        Return 0
    End Function

    Private Function _GetGinkouSaffix(ByRef koumei As String) As String
        If koumei.Contains("銀行") Then
            Return "銀行"
        ElseIf koumei.Contains("信用金庫") Then
            Return "信用金庫"
        ElseIf koumei.Contains("信用組合") Then
            Return "信用組合"
        ElseIf koumei.Contains("農業協同組") Then
            Return "農業協同組合"
        ElseIf koumei.Contains("農協") Then
            Return "農協"
        Else
            Return vbNullString
        End If
    End Function

    Private Function _GetGinkou(ByVal koumei As String) As String
        koumei.Replace("銀行", vbNullString)
        koumei.Replace("信用金庫", vbNullString)
        koumei.Replace("信用組合", vbNullString)
        koumei.Replace("農業協同組合", vbNullString)
        koumei.Replace("農協", vbNullString)
        Return koumei
    End Function

    Private Function _GetTuchoKigouBangou(ByVal kouza As String) As String
        Dim bno As String
        Dim kno As String
        bno = Strings.Right(kouza, 7)
        kno = Strings.Replace(kouza, bno, vbNullString,,, CompareMethod.Text)
        bno &= "1"
        kno = "1" & kno.PadLeft(4, "0")
        Return (kno & bno)
    End Function


    Private Sub _CreateSummaryFile(ByVal otfile As String)
        Dim sw As StreamWriter
        Dim bankcount = 0
        Dim sumcount = 0
        Dim cimds As List(Of IraishoMeisai)
        Dim wline = vbNullString
        Dim zimd = New IraishoMeisai

        sw = New StreamWriter(otfile, False, System.Text.Encoding.GetEncoding(SHIFT_JIS))

        For Each imd In Me.IraishoMeisaiDatas
            If imd.T_Ginkoucode <> zimd.T_Ginkoucode Then
                cimds = Me.IraishoMeisaiDatas.FindAll(
                    Function(cimd)
                        Return (cimd.T_Ginkoucode = imd.T_Ginkoucode)
                    End Function
                )
                bankcount = cimds.Count
                wline = imd.T_Ginkoumei.PadRight(18, "　")
                wline &= bankcount.ToString().PadLeft(3, " ")
                sw.WriteLine(wline)
                wline = vbNullString
                zimd = imd
            End If
        Next

        ' 合計
        sumcount = Me.IraishoMeisaiDatas.Count
        wline = "合計件数".PadRight(18, "　")
        wline &= sumcount.ToString().PadLeft(3, " ")
        sw.WriteLine(wline)
        wline = vbNullString

        If sw IsNot Nothing Then
            sw.Close()
            sw.Dispose()
        End If
    End Sub


    Private Sub _CreateIraishoMeisaiDatas(ByVal infile As String, ByVal otfile As String)
        Dim sr As StreamReader
        Dim sw As StreamWriter

        Dim vs1(), vs2()                                     ' ＣＳＶ各要素
        Dim line As String                                   ' 入力テキスト

        Dim imds = New List(Of IraishoMeisai)
        Dim simd As IraishoMeisai = Nothing
        Dim zimd As IraishoMeisai = Nothing

        Dim R_obj = Rpa.RootProjectObject
        Dim R_iraishos = CType(R_obj.IraishoDatas, List(Of Iraisho))
        Dim R_iraisho As Iraisho = Nothing
        Dim R_tb As Iraisho.BankInfo = Nothing

        Dim dt As DateTime                                   ' 日付関係
        Dim culture As CultureInfo
        Dim sYYYYMMDD, sMM, sDD, wYYYYMMDD, wYY

        Dim header                                           ' ヘッダーフラグ
        Dim resheet                                          ' 改ページフラグ
        Dim row                                              ' 行位置
        Dim srcsheetname                                     ' 作成元シート名
        Dim dstsheetname                                     ' 作成先シート名
        Dim tno, tkno, tbno                                  ' 通帳記号番号関連
        Dim zsc                                              ' 自営信金コード
        Dim gc, gm, sm                                       ' 銀行コード、銀行名、支店名

        Dim imdcount = 0                                     ' 依頼書の件数

        Dim wline As String                                  ' 出力テキスト


        ' ストリームリーダ、ライター
        sr = New StreamReader(infile, System.Text.Encoding.GetEncoding(MyEncoding))
        sw = New StreamWriter(otfile, False, System.Text.Encoding.GetEncoding(SHIFT_JIS))

        ' 今日の日付
        dt = DateTime.Now
        sYYYYMMDD = dt.ToString("yyyyMMdd")
        sMM = Strings.Mid(sYYYYMMDD, 5, 2)
        sDD = Strings.Mid(sYYYYMMDD, 7, 2)
        culture = New CultureInfo("ja-JP", True)
        culture.DateTimeFormat.Calendar = New JapaneseCalendar
        wYYYYMMDD = dt.ToString("ggyyMMdd", culture)

        ' ヘッダーは読み飛ばし
        sr.ReadLine()

        ' MeisaiStringはソート後に再度Splitし、並べ替えを行う
        Do Until sr.EndOfStream
            line = sr.ReadLine()
            vs1 = line.Split(",")
            simd = New IraishoMeisai With {
                .MeisaiString = line,
                .T_Ginkoucode = vs1(7),
                .Sitencode = vs1(8),
                .T_Ginkoumei = vs1(9)
            }

            imds.Add(simd)
        Loop

        ' 銀行コード－＞支店コードでソート
        imds.Sort(
            Function(a, b)
                If a.T_Ginkoucode <> b.T_Ginkoucode Then
                    Return (a.T_Ginkoucode - b.T_Ginkoucode)
                Else
                    Return (a.Sitencode - b.Sitencode)
                End If
            End Function
        )

        zimd = New IraishoMeisai
        For Each imd In imds
            header = "0"

            ' 銀行が変わったとき
            If zimd.T_Ginkoucode <> imd.T_Ginkoucode Then
                imdcount = 0
                R_iraisho = Nothing
                R_iraisho = R_iraishos.Find(                                   '1. 依頼書の対象銀行コードが合致する依頼書データ取得
                    Function(ri)
                        Dim hit = ri.TargetBanks.Find(
                            Function(tb)
                                Return (tb.Code = imd.T_Ginkoucode)
                            End Function
                        )
                        Return IIf(hit IsNot Nothing, True, False)
                    End Function
                )
                If R_iraisho Is Nothing Then                                   '2. 対象パターン文字列群から依頼書データを取得
                    R_iraisho = R_iraishos.Find(
                        Function(ri)
                            If ri.TargetPattern IsNot Nothing Then
                                Dim hit = ri.TargetPattern.Find(
                                    Function(tp)
                                        Return imd.T_Ginkoumei.Contains(tp)
                                    End Function
                                )
                                Return IIf(hit IsNot Nothing, True, False)
                            Else
                                Return False
                            End If
                        End Function
                    )
                End If
                If R_iraisho Is Nothing Then                                   '3. デフォルト依頼書データを取得
                    R_iraisho = R_iraishos(0)
                End If

                ' 銀行情報取得（なければ空の銀行情報をセット）
                R_tb = Nothing
                R_tb = R_iraisho.TargetBanks.Find(
                    Function(tb)
                        Return (tb.Code = imd.T_Ginkoucode)
                    End Function
                )

                header = "1"
                dstsheetname = vbNullString
            End If

            imdcount += 1

            ' 改ページ判定
            resheet = IIf(header = "1", "1", "0")
            Select Case R_iraisho.ResheetType
                Case "WhenReachMaxCount"
                    resheet = IIf(imdcount > R_iraisho.MaxCount, "1", resheet)
                Case "WhenReachMaxCountAndSitenChanged"
                    resheet = IIf(imdcount > R_iraisho.MaxCount, "1", resheet)
                    resheet = IIf(imd.Sitencode <> zimd.Sitencode, "1", resheet)
            End Select
            imdcount = IIf(resheet = "1", 1, imdcount)

            ' 作成元シート名
            srcsheetname = R_iraisho.SourceSheetName
            If R_tb IsNot Nothing Then
                If Not String.IsNullOrEmpty(R_tb.SourceSheetName) Then
                    srcsheetname = R_tb.SourceSheetName
                End If
            End If

            ' 作成先シート名
            If resheet = "1" Then
                If String.IsNullOrEmpty(dstsheetname) Then
                    dstsheetname = imd.T_Ginkoumei
                Else
                    dstsheetname &= "i"
                End If
            End If

            ' 行数
            row = (R_iraisho.MeisaiTableTopRow - 1 + imdcount).ToString()

            vs2 = imd.MeisaiString.Split(",")

            ' 加工データ
            '---------------------------------------------------------------------------'
            gc = vs2(7).PadLeft(4, "0")
            If R_tb Is Nothing Then
                gm = vs2(9)
                zsc = vbNullString
            Else
                gm = R_tb.Name
                zsc = R_tb.ZieiSinkinCode
            End If
            sm = vs2(10)
            If gc = "9900" Then
                tno = _GetTuchoKigouBangou(vs2(12))
                tkno = Strings.Left(tno, 5)
                tbno = Strings.Right(tno, 8)
            Else
                tkno = vbNullString
                tbno = vbNullString
            End If
            '---------------------------------------------------------------------------'

            ' 未使用
            '---------------------------------------------------------------------------'
            R_iraisho.FurikaesiteibiCell = vbNullString
            '---------------------------------------------------------------------------'

            ' 各種依頼書情報付加
            wline &= $"{Me.Work2Directory}\{R_iraisho.BookName}"                                                                            ' idx000 Ｅｘｃｅｌ名
            wline &= "," & srcsheetname                                                                                                     ' idx001
            wline &= "," & dstsheetname                                                                                                     ' idx002
            wline &= "," & R_obj.Syunoukigyoucode                                                                                           ' idx003
            wline &= "," & R_iraisho.SyunoukigyoucodeCell                                                                                   ' idx004
            wline &= "," & R_obj.Syunoukigyoumei                                                                                            ' idx005
            wline &= "," & R_iraisho.SyunoukigyoumeiCell                                                                                    ' idx006
            wline &= "," & R_obj.Syunoukigyouaddress                                                                                        ' idx007
            wline &= "," & R_iraisho.SyunoukigyouaddressCell                                                                                ' idx008
            wline &= "," & R_obj.Syunoukigyoutel                                                                                            ' idx009
            wline &= "," & R_iraisho.SyunoukigyoutelCell                                                                                    ' idx010
            wline &= "," & R_obj.Syunoukigyoutantousya                                                                                      ' idx011
            wline &= "," & R_iraisho.SyunoukigyoutantousyaCell                                                                              ' idx012
            wline &= "," & R_obj.Haraikomisakikouzabangou                                                                                   ' idx013
            wline &= "," & R_iraisho.HaraikomisakikouzabangouCell                                                                           ' idx014
            wline &= "," & R_obj.Haraikomikinshubetu                                                                                        ' idx015
            wline &= "," & R_iraisho.HaraikomikinshubetuCell                                                                                ' idx016
            wline &= "," & sYYYYMMDD                                                                                                        ' idx017 依頼日
            wline &= "," & R_iraisho.IraibiCell                                                                                             ' idx018
            wline &= "," & sYYYYMMDD                                                                                                        ' idx019 作成日
            wline &= "," & R_iraisho.SakuseibiCell                                                                                          ' idx020
            wline &= "," & Strings.Left(sYYYYMMDD, 4)                                                                                       ' idx021 作成年（４桁）
            wline &= "," & R_iraisho.SakuseibiYYYYCell                                                                                      ' idx022
            wline &= "," & Strings.Mid(sYYYYMMDD, 3, 2)                                                                                     ' idx023 作成年（下２桁）
            wline &= "," & R_iraisho.SakuseibiYYCell                                                                                        ' idx024
            wline &= "," & wYY                                                                                                              ' idx025 作成年（和暦）
            wline &= "," & R_iraisho.SakuseibiWaYYCell                                                                                      ' idx026
            wline &= "," & IIf(wYY > "09", wYY, Strings.Right(wYY, 1))                                                                      ' idx027 作成年（和暦、ゼロサプレス）
            wline &= "," & R_iraisho.SakuseibiWaZYCell                                                                                      ' idx028
            wline &= "," & sMM                                                                                                              ' idx029 作成月
            wline &= "," & R_iraisho.SakuseibiMMCell                                                                                        ' idx030
            wline &= "," & IIf(sMM > "09", sMM, Strings.Right(sMM, 1))                                                                      ' idx031 作成月（ゼロサプレス）
            wline &= "," & R_iraisho.SakuseibiZMCell                                                                                        ' idx032
            wline &= "," & sDD                                                                                                              ' idx033 作成日
            wline &= "," & R_iraisho.SakuseibiDDCell                                                                                        ' idx034
            wline &= "," & IIf(sDD > "09", sDD, Strings.Right(sDD, 1))                                                                      ' idx035 作成日（ゼロサプレス）
            wline &= "," & R_iraisho.SakuseibiZDCell                                                                                        ' idx036
            wline &= "," & vbNullString                                                                                                     ' idx037 振替指定日
            wline &= "," & R_iraisho.FurikaesiteibiCell                                                                                     ' idx038 （現在未使用）
            wline &= "," & vs2(2)                                                                                                           ' idx039 委託者コード
            wline &= "," & R_iraisho.ItakusyacodeCell                                                                                       ' idx040
            wline &= "," & vs2(3)                                                                                                           ' idx041 委託者名
            wline &= "," & R_iraisho.ItakusyameiCell                                                                                        ' idx042
            wline &= "," & R_iraisho.Iraisaki                                                                                               ' idx043 依頼先
            wline &= "," & R_iraisho.IraisakiCell                                                                                           ' idx044
            wline &= "," & R_iraisho.I_Ginkoucode                                                                                           ' idx045 銀行コード
            wline &= "," & R_iraisho.I_GinkoucodeCell                                                                                       ' idx046
            wline &= "," & R_iraisho.I_Ginkoucode1                                                                                          ' idx047 銀行コード（左から１桁目）
            wline &= "," & R_iraisho.I_Ginkoucode1Cell                                                                                      ' idx048
            wline &= "," & R_iraisho.I_Ginkoucode2                                                                                          ' idx049 銀行コード（左から２桁目）
            wline &= "," & R_iraisho.I_Ginkoucode2Cell                                                                                      ' idx050
            wline &= "," & R_iraisho.I_Ginkoucode3                                                                                          ' idx051 銀行コード（左から３桁目）
            wline &= "," & R_iraisho.I_Ginkoucode3Cell                                                                                      ' idx052
            wline &= "," & R_iraisho.I_Ginkoucode4                                                                                          ' idx053 銀行コード（左から４桁目）
            wline &= "," & R_iraisho.I_Ginkoucode4Cell                                                                                      ' idx054
            wline &= "," & R_iraisho.I_Ginkoumei                                                                                            ' idx055 銀行名
            wline &= "," & R_iraisho.I_GinkoumeiCell                                                                                        ' idx056
            wline &= "," & R_iraisho.I_Ginkou                                                                                               ' idx057 銀行名（銀行名から「銀行」などを除いた名称）
            wline &= "," & R_iraisho.I_GinkouCell                                                                                           ' idx058
            wline &= "," & R_iraisho.I_GinkouSaffix                                                                                         ' idx059 「銀行」・「信用金庫」など
            wline &= "," & R_iraisho.I_GinkouSaffixCell                                                                                     ' idx060
            wline &= "," & R_iraisho.I_Zieisinkincode                                                                                       ' idx061 自営信金コード
            wline &= "," & R_iraisho.I_ZieisinkincodeCell                                                                                   ' idx062
            wline &= "," & R_iraisho.Baitaimei                                                                                              ' idx063 媒体名
            wline &= "," & R_iraisho.BaitaimeiCell                                                                                          ' idx064
            wline &= "," & gc                                                                                                               ' idx065 銀行コード
            wline &= "," & R_iraisho.T_GinkoucodeCell                                                                                       ' idx066
            wline &= "," & gc(0).ToString()                                                                                                 ' idx067 銀行コード（左から１桁目）
            wline &= "," & R_iraisho.T_Ginkoucode1Cell                                                                                      ' idx068
            wline &= "," & gc(1).ToString()                                                                                                 ' idx069 銀行コード（左から２桁目）
            wline &= "," & R_iraisho.T_Ginkoucode2Cell                                                                                      ' idx070
            wline &= "," & gc(2).ToString()                                                                                                 ' idx071 銀行コード（左から３桁目）
            wline &= "," & R_iraisho.T_Ginkoucode3Cell                                                                                      ' idx072
            wline &= "," & gc(3).ToString()                                                                                                 ' idx073 銀行コード（左から４桁目）
            wline &= "," & R_iraisho.T_Ginkoucode4Cell                                                                                      ' idx074
            wline &= "," & gm                                                                                                               ' idx075 銀行名
            wline &= "," & R_iraisho.T_GinkoumeiCell                                                                                        ' idx076
            wline &= "," & _GetGinkou(gm)                                                                                                   ' idx077 銀行名（銀行名から「銀行」などを除いた名称）
            wline &= "," & R_iraisho.T_GinkouCell                                                                                           ' idx078
            wline &= "," & _GetGinkouSaffix(gm)                                                                                             ' idx079 「銀行」・「信用金庫」など
            wline &= "," & R_iraisho.T_GinkouSaffixCell                                                                                     ' idx080
            wline &= "," & zsc                                                                                                              ' idx081 自営信金コード
            wline &= "," & R_iraisho.T_ZieisinkincodeCell                                                                                   ' idx082
            wline &= "," & vs2(4)                                                                                                           ' idx083 顧客コード
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.KokyakucodeColumn), vbNullString, R_iraisho.KokyakucodeColumn & row)          ' idx084
            wline &= "," & gc                                                                                                               ' idx085 銀行コード
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.K_GinkoucodeColumn), vbNullString, R_iraisho.K_GinkoucodeColumn & row)        ' idx086
            wline &= "," & vs2(8)                                                                                                           ' idx087 支店コード
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.SitencodeColumn), vbNullString, R_iraisho.SitencodeColumn & row)              ' idx088
            wline &= "," & gm                                                                                                               ' idx089 銀行名
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.K_GinkoumeiColumn), vbNullString, R_iraisho.K_GinkoumeiColumn & row)          ' idx090
            wline &= "," & _GetGinkou(gm)                                                                                                   ' idx091 銀行名（銀行名から「銀行」などを除いた名称）
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.K_GinkouColumn), vbNullString, R_iraisho.K_GinkouColumn & row)                ' idx092
            wline &= "," & sm                                                                                                               ' idx093 支店名
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.SitenmeiColumn), vbNullString, R_iraisho.SitenmeiColumn & row)                ' idx094
            wline &= "," & sm.Replace("支店", vbNullString)                                                                                 ' idx095 支店名（支店名から「支店」を除いた名称）
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.SitenColumn), vbNullString, R_iraisho.SitenColumn & row)                      ' idx096
            wline &= "," & vs2(11)                                                                                                          ' idx097 預金種別
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.YokinshubetuColumn), vbNullString, R_iraisho.YokinshubetuColumn & row)        ' idx098
            wline &= "," & vs2(12)                                                                                                          ' idx099 口座番号
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.KouzabangouColumn), vbNullString, R_iraisho.KouzabangouColumn & row)          ' idx100
            wline &= "," & tkno                                                                                                             ' idx101 通帳記号
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.TuchoukigouColumn), vbNullString, R_iraisho.TuchoukigouColumn & row)          ' idx102
            wline &= "," & tbno                                                                                                             ' idx103 通帳番号
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.TuchoubangouColumn), vbNullString, R_iraisho.TuchoubangouColumn & row)        ' idx104
            wline &= "," & vs2(13)                                                                                                          ' idx105 口座名義（半角）
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.KouzameigiColumn), vbNullString, R_iraisho.KouzameigiColumn & row)            ' idx106
            wline &= "," & Strings.StrConv(vs2(13), VbStrConv.Wide)                                                                         ' idx107 口座名義（全角）
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.JISKouzameigiColumn), vbNullString, R_iraisho.JISKouzameigiColumn & row)      ' idx108
            wline &= "," & vs2(14)                                                                                                          ' idx109 振替金額
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.FurikaekingakuColumn), vbNullString, R_iraisho.FurikaekingakuColumn & row)    ' idx110
            wline &= "," & R_iraisho.Bikou                                                                                                  ' idx111 備考
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.BikouColumn), vbNullString, R_iraisho.BikouColumn & row)                      ' idx112
            wline &= "," & R_iraisho.Tekiyou                                                                                                ' idx113 適用
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.TekiyouColumn), vbNullString, R_iraisho.TekiyouColumn & row)                  ' idx114
            wline &= "," & R_iraisho.Funoucode                                                                                              ' idx115 不能コード
            wline &= "," & IIf(String.IsNullOrEmpty(R_iraisho.FunoucodeColumn), vbNullString, R_iraisho.FunoucodeColumn & row)              ' idx116
            wline &= "," & vs2(19)                                                                                                          ' idx117 エラーコード

            sw.WriteLine(wline)
            wline = vbNullString
            zimd = imd
        Next

        ' 件数集計用に保持
        Me.IraishoMeisaiDatas = imds

        If sr IsNot Nothing Then
            sr.Close()
            sr.Dispose()
        End If
        If sw IsNot Nothing Then
            sw.Close()
            sw.Dispose()
        End If
    End Sub


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

            ' リスタートの設定
            Me.RestartCode = iv(4)

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
