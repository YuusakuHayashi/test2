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

    Private _Syunoukigyoucode As String
    Public Property Syunokigyoucode As String              ' 収納企業コード
        Get
            Return Me._Syunoukigyoucode
        End Get
        Set(value As String)
            Me._Syunoukigyoucode = value
        End Set
    End Property

    Private _Syunoukigyoumei As String
    Public Property Syunokigyoumei As String               ' 収納企業名
        Get
            Return Me._Syunoukigyoumei
        End Get
        Set(value As String)
            Me._Syunoukigyoumei = value
        End Set
    End Property

    Private _Syunoukigyouaddress As String
    Public Property Syunokigyouaddress As String           ' 収納企業住所
        Get
            Return Me._Syunoukigyouaddress
        End Get
        Set(value As String)
            Me._Syunoukigyouaddress = value
        End Set
    End Property

    Private _Syunoukigyoutel As String
    Public Property Syunokigyoutel As String               ' 収納企業電話番号
        Get
            Return Me._Syunoukigyoutel
        End Get
        Set(value As String)
            Me._Syunoukigyoutel = value
        End Set
    End Property

    Private _Syunoukigyoutantousya As String
    Public Property Syunokigyoutantousya As String         ' 収納企業担当者
        Get
            Return Me._Syunoukigyoutantousya
        End Get
        Set(value As String)
            Me._Syunoukigyoutantousya = value
        End Set
    End Property

    Private _Haraikomisakikouzabangou As String
    Public Property Haraikomisakikouzabangou As String     ' 払込先口座番号
        Get
            Return Me._Haraikomisakikouzabangou
        End Get
        Set(value As String)
            Me._Haraikomisakikouzabangou = value
        End Set
    End Property

    Private _Haraikomikinshubetu As String
    Public Property Haraikomikinshubetu As String          ' 払込金種別
        Get
            Return Me._Haraikomikinshubetu
        End Get
        Set(value As String)
            Me._Haraikomikinshubetu = value
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
        Public SheetName As String
        Public MaxCount As Integer
        Public TargetPattern As List(Of String)
        Public ResheetType As String

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
        Public KokyakucodeColumn As String                ' 顧客コード
        Public K_GinkoucodeColumn As String               ' 銀行コード
        Public SitencodeColumn As String                  ' 支店コード
        Public K_GinkoumeiColumn As String                ' 銀行名
        Public K_GinkouColumn As String                   ' 銀行名（銀行名から「銀行」などを除いた名称）
        Public SitenmeiColumn As String                   ' 支店名
        Public SitenColumn As String                      ' 支店名（支店名から「支店」を除いた名称）
        Public YokinshubetuColumn As String               ' 預金種別
        Public KouzabangoColumn As String                 ' 口座番号
        Public TuchoukigouColumn As String                ' 通帳記号
        Public TuchoubangouColumn As String               ' 通帳番号
        Public KouzameigiColumn As String                 ' 口座名義
        Public FurikaekingakuColumn As String             ' 振替金額
        Public ErrorCodeColumn As String                  ' エラーコード（入力情報との突合せ結果）

        Public BikouColumn As String                      ' 備考
        Public TekiyouColumn As String                    ' 摘要
        Public FunoucodeColumn As String                  ' 不能コード
        Public Bikou As String                            ' 備考データ
        Public Tekiyou As String                          ' 摘要データ
        Public Funoucode As String                        ' 不能コードデータ

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
    Public Class IraishoMeisai
        ' ＭＢＣ
        Public Syunokigyoucode As String              ' 収納企業コード
        Public Syunoukigyoumei As String              ' 収納企業名
        Public Syunoukigyouaddress As String          ' 収納企業住所
        Public Syunoukigyoutel As String              ' 収納企業電話番号
        Public Syunoukigyoutantousya As String        ' 収納企業担当者
        Public Haraikomisakikouzabangou As String     ' 払込先口座番号
        Public Haraikomikinshubetu As String          ' 払込金種別
        Public Iraibi As String                       ' 依頼日
        Public Sakuseibi As String                    ' 作成日（８桁）
        Public SakuseibiYYYY As String                ' 作成年（４桁）
        Public SakuseibiYY As String                  ' 作成年（下２桁）
        Public SakuseibiWaYY As String                ' 作成年（和暦）
        Public SakuseibiMM As String                  ' 作成月（２桁）
        Public SakuseibiZM As String                  ' 作成月（９月以前は１桁）
        Public SakuseibiDD As String                  ' 作成日（２桁）
        Public SakuseibiZD As String                  ' 作成日（９日以前は１桁）
        Public Furikaesiteibi As String               ' 振替指定日

        ' 停止依頼元
        Public Itakusyacode As String                 ' 委託者コード
        Public Itakusyamei As String                  ' 委託者名

        ' 停止依頼先
        Public Iraisaki As String                     ' 提出先
        Public I_Ginkoucode As String                 ' 銀行コード
        Public I_Ginkoucode1 As String                ' 銀行コード（左から１桁目）
        Public I_Ginkoucode2 As String                ' 銀行コード（左から２桁目）
        Public I_Ginkoucode3 As String                ' 銀行コード（左から３桁目）
        Public I_Ginkoucode4 As String                ' 銀行コード（左から４桁目）
        Public I_Ginkoumei As String                  ' 銀行名
        Public I_Ginkou As String                     ' 銀行名（銀行名から「銀行」などを除いた名称）
        Public I_GinkouSaffix As String               ' 「銀行」・「信用金庫」など
        Public I_Zieisinkincode As String             ' 自営信金コード
        Public Baitaimei As String                    ' 媒体名

        ' 停止銀行
        Public T_Ginkoucode As String                 ' 銀行コード
        Public T_Ginkoucode1 As String                ' 銀行コード（左から１桁目）
        Public T_Ginkoucode2 As String                ' 銀行コード（左から２桁目）
        Public T_Ginkoucode3 As String                ' 銀行コード（左から３桁目）
        Public T_Ginkoucode4 As String                ' 銀行コード（左から４桁目）
        Public T_Ginkoumei As String                  ' 銀行名
        Public T_Ginkou As String                     ' 銀行名（銀行名から「銀行」などを除いた名称）
        Public T_GinkouSaffix As String               ' 「銀行」・「信用金庫」など
        Public T_Zieisinkincode As String             ' 自営信金コード

        ' 停止対象
        Public Kokyakucode As String                  ' 顧客コード
        Public Sitencode As String                    ' 支店コード
        Public Sitenmei As String                     ' 支店名
        Public Siten As String                        ' 支店名（支店名から「支店」を除いた名称）
        Public Yokinshubetu As String                 ' 預金種別
        Public Kouzabango As String                   ' 口座番号
        Public Tuchoukigou As String                  ' 通帳記号
        Public Tuchoubangou As String                 ' 通帳番号
        Public Kouzameigi As String                   ' 口座名義
        Public Furikaekingaku As String               ' 振替金額
        Public ErrorCode As String                    ' エラーコード（入力情報との突合せ結果）
        Public Bikou As String                        ' 備考
        Public Tekiyou As String                      ' 摘要
        Public Funoucode As String                     ' 不能コード

        ' その他
        Public HeaderFlag As String
        Public ResheetFlag As String
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
        Dim inmaster2 = inmaster
        Dim tmpcsv = Rpa.MyProjectWorkDirectory & "\tmp.csv"
        Dim target = Me.ItakusyaCodeDictionary("Moteki")
        Dim incsv2 = incsv
        Dim tincsv = Rpa.MyProjectWorkDirectory & "\t_input.csv"
        Dim tincsv2 = Rpa.MyProjectWorkDirectory & "\t_input2.csv"
        Call _CreateTmpCsv(inmaster2, target, tmpcsv)
        Call _CompareInputCsvToTmpCsv(incsv, tmpcsv, tincsv)

        Call _CreateIraishoMeisaiDatas(tincsv, tincsv2)
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

    Private Function _SetGinkoucode(ByVal bankcode As String)
        Dim bc = meisai.T_Ginkoucode.PadLeft(4, "0")
        meisai.T_Ginkoucode1 = meisai.T_Ginkoucode(0).ToString
        meisai.T_Ginkoucode2 = meisai.T_Ginkoucode(1).ToString
        meisai.T_Ginkoucode3 = meisai.T_Ginkoucode(2).ToString
        meisai.T_Ginkoucode4 = meisai.T_Ginkoucode(3).ToString
    End Function

    Private Function _GetGinkouSaffix(ByRef koumei As String) As String
        If koumei.Contains("銀行") Then
            Return "銀行"
        ElseIf koumei.Contains("信用金庫") Then
            Return "信用金庫"
        ElseIf koumei.Contains("信用組合") Then
            Return "信用組合"
        Else
            Return vbNullString
        End If
    End Function

    Private Function _GetGinkou(ByVal koumei As String) As String
        koumei.Replace("銀行", vbNullString)
        koumei.Replace("信用金庫", vbNullString)
        koumei.Replace("信用組合", vbNullString)
        koumei.Replace("農業協同組合", vbNullString)
        Return koumei
    End Function

    Private Sub _SetSiten(ByRef meisai As IraishoMeisai)
        meisai.Siten = meisai.Sitenmei.Replace("支店", vbNullString)
    End Sub

    Private Sub _SetTuchoKigouBangou(ByRef meisai As IraishoMeisai)
        Dim no As String
        Dim bno As String
        Dim kno As String
        If meisai.T_Ginkoucode = "9900" Then
            no = meisai.Kouzabango
            bno = Strings.Right(no, 7)
            kno = Strings.Replace(no, bno, vbNullString,,, CompareMethod.Text)
            bno &= "1"
            kno = "1" & kno.PadLeft(4, "0")
            meisai.Tuchoukigou = kno
            meisai.Tuchoubangou = bno
        Else
            meisai.Tuchoukigou = vbNullString
            meisai.Tuchoubangou = vbNullString
        End If
    End Sub

    Private Sub _CreateIraishoMeisaiDatas(ByVal infile As String, ByVal otfile As String)
        Dim slist As New List(Of String)
        Dim sr As StreamReader
        Dim sw As StreamWriter
        Dim vs1() As String
        Dim vs2() As String
        Dim line As String
        Dim wline As String
        Dim gc As String
        Dim meisai As IraishoMeisai
        Dim zimd As IraishoMeisai
        Dim miraisho As Iraisho
        Dim R_iraishos = CType(Rpa.RootProjectObject.IraishoDatas, List(Of Iraisho))
        Dim dt = DateTime.Now
        Dim sYYYYMMDD = dt.ToString("yyyyMMdd")
        Dim sMM = Strings.Mid(sYYYYMMDD, 5, 2)
        Dim sDD = Strings.Mid(sYYYYMMDD, 7, 2)
        Dim culture = New CultureInfo("ja-JP", True)
        culture.DateTimeFormat.Calender = New JapaneseCalender()
        Dim wYYYYMMDD = dt.ToString("ggyyMMdd", culture)
        Dim wYY = Strings.Mid(wYYYYMMDD, 3, 2)
        sr = New StreamReader(infile, System.Text.Encoding.GetEncoding(MyEncoding))
        sw = New StreamWriter(otfile, False, System.Text.Encoding.GetEncoding(MyEncoding))


        ' ヘッダーは読み飛ばし
        sr.ReadLine()

        ' MeisaiStringはソート後に再度Splitし、並べ替えを行う
        Do Until sr.EndOfStream
            line = sr.ReadLine()
            vs1 = line.Split(",")
            meisai = New IraishoMeisai With {
                .MeisaiString = line,
                .Itakusyacode = vs1(2),
                .Itakusyamei = vs1(3),
                .Kokyakucode = vs1(4),
                .T_Ginkoucode = vs1(7),
                .Sitencode = vs1(8),
                .T_Ginkoumei = vs1(9),
                .Sitenmei = vs1(10),
                .Yokinshubetu = vs1(11),
                .Kouzabango = vs1(12),
                .Kouzameigi = vs1(13),
                .Furikaekingaku = vs1(14),
                .ErrorCode = vs1(19)
            }

            Call _SetGinkou(meisai)
            Call _SetGinkouSaffix(meisai)
            Call _SetGinkoucode(meisai)
            Call _SetSiten(meisai)
            Call _SetTuchoKigouBangou(meisai)

            Me.IraishoMeisaiDatas.Add(meisai)
        Loop

        ' 銀行コード－＞支店コードでソート
        Me.IraishoMeisaiDatas.Sort(
            Function(a, b)
                If a.T_Ginkoucode <> b.T_Ginkoucode Then
                    Return (a.T_Ginkoucode - b.T_Ginkoucode)
                Else
                    Return (a.Sitencode - b.Sitencode)
                End If
            End Function
        )

        Dim imdcount = 0
        zimd = New IraishoMeisai
        For Each imd In Me.IraishoMeisaiDatas

            ' 銀行が変わったとき
            ' 1. 依頼書の対象銀行コードが合致する依頼書データを取得
            ' 2. 対象パターン文字列群から依頼書データを取得
            ' 3. 上記に該当しない場合、デフォルト依頼書データを取得
            If zimd.T_Ginkoucode <> imd.T_Ginkoucode Then
                imdcount = 0
                miraisho = Nothing
                miraisho = R_iraishos.Find(                                   '1.
                    Function(ri)
                        Dim hit = ri.TargetBanks.Find(
                            Function(tb)
                                Return (tb.Code = imd.T_Ginkoucode)
                            End Function
                        )
                        Return IIf(hit IsNot Nothing, True, False)
                    End Function
                )
                If miraisho Is Nothing Then                                   '2.
                    miraisho = R_iraishos.Find(
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
                If miraisho Is Nothing Then                                   '3.
                    miraisho = R_iraishos(0)
                End If

                imd.HeaderFlag = True
            End If

            vs2 = imd.MeisaiString.Split(",")

            imd.I_Ginkoucode = miraisho.I_Ginkoucode
            imd.I_Ginkoucode1 = miraisho.I_Ginkoucode1
            imd.I_Ginkoucode2 = miraisho.I_Ginkoucode2
            imd.I_Ginkoucode3 = miraisho.I_Ginkoucode3
            imd.I_Ginkoucode4 = miraisho.I_Ginkoucode4
            imd.I_Ginkoumei = miraisho.I_Ginkoumei
            imd.I_Ginkou = miraisho.I_Ginkou
            imd.I_GinkouSaffix = miraisho.I_GinkouSaffix
            imd.I_Zieisinkincode = miraisho.I_Zieisinkincode
            imd.Baitaimei = miraisho.Baitaimei
            imd.Bikou = miraisho.Bikou
            imd.Tekiyou = miraisho.Tekiyou
            imd.Funoucode = miraisho.Funoucode

            imdcount += 1
            imd.ResheetFlag = IIf(imdcount > miraisho.MaxCount, True, False)
            imdcount = IIf(imd.ResheetFlag, 1, imdcount)

            ' 各種依頼書情報付加
            wline &= miraisho.BookName                                    ' idx21
            wline &= "," & miraisho.SheetName                             ' idx22
            wline &= "," & miraisho.ResheetType                           ' idx23
            wline &= "," & Rpa.RootProjectObject.Syunoukigyoucode         ' idx23
            wline &= "," & miraisho.SyunoukigyoucodeCell                  ' idx23
            wline &= "," & Rpa.RootProjectObject.Syunoukigyoumei          ' idx23
            wline &= "," & miraisho.SyunoukigyoumeiCell                   ' idx23
            wline &= "," & Rpa.RootProjectObject.Syunoukigyouaddress      ' idx23
            wline &= "," & miraisho.SyunoukigyouaddressCell               ' idx23
            wline &= "," & Rpa.RootProjectObject.Syunoukigyouaddress      ' idx23
            wline &= "," & miraisho.SyunoukigyouaddressCell               ' idx23
            wline &= "," & Rpa.RootProjectObject.Syunoukigyoutel          ' idx23
            wline &= "," & miraisho.SyunoukigyoutelCell                   ' idx23
            wline &= "," & Rpa.RootProjectObject.Syunoukigyoutantousya    ' idx23
            wline &= "," & miraisho.SyunoukigyoutantousyaCell             ' idx23
            wline &= "," & Rpa.RootProjectObject.Haraikomisakikouzabangou ' idx23
            wline &= "," & miraisho.HaraikomisakikouzabangouCell          ' idx23
            wline &= "," & Rpa.RootProjectObject.Haraikomikishubetu       ' idx23
            wline &= "," & miraisho.HaraikomikishubetuCell                ' idx23
            wline &= "," & sYYYYMMDD                                      ' idx23 依頼日
            wline &= "," & miraisho.IraibiCell                            ' idx23
            wline &= "," & sYYYYMMDD                                      ' idx23 作成日
            wline &= "," & miraisho.SakuseibiCell                         ' idx23
            wline &= "," & Strings.Left(sYYYYMMDD, 4)                     ' idx23 作成年（４桁）
            wline &= "," & miraisho.SakuseibiYYYYCell                     ' idx23
            wline &= "," & Strings.Mid(sYYYYMMDD, 3, 2)                   ' idx23 作成年（下２桁）
            wline &= "," & miraisho.SakuseibiYYCell                       ' idx23
            wline &= "," & wYY                                            ' idx23 作成年（和暦）
            wline &= "," & miraisho.SakuseibiWaYYCell                     ' idx23
            wline &= "," & IIf(wYY > "09", wYY, Strings.Right(wYY, 1))    ' idx23 作成年（和暦、ゼロサプレス）
            wline &= "," & miraisho.SakuseibiWaZYCell                     ' idx23
            wline &= "," & sMM                                            ' idx23 作成月
            wline &= "," & miraisho.SakuseibiMMCell                       ' idx23
            wline &= "," & IIf(sMM > "09", sMM, Strings.Right(sMM, 1))    ' idx23 作成月（ゼロサプレス）
            wline &= "," & miraisho.SakuseibiZMCell                       ' idx23
            wline &= "," & sDD                                            ' idx23 作成日
            wline &= "," & miraisho.SakuseibiDDCell                       ' idx23
            wline &= "," & IIf(sDD > "09", sDD, Strings.Right(sDD, 1))    ' idx23 作成日（ゼロサプレス）
            wline &= "," & miraisho.SakuseibiZDCell                       ' idx23

            ' 振替指定日は現在未使用
            '-------------------------------------------------------------------------
            miraisho.FurikaesiteibiCell = vbNullString
            wline &= "," & vbNullString                                   ' idx23 振替指定日
            wline &= "," & miraisho.FurikaesiteibiCell                    ' idx23
            '-------------------------------------------------------------------------

            wline &= "," & vs2(2)                                         ' idx23 委託者コード
            wline &= "," & miraisho.ItakusyacodeCell                      ' idx23
            wline &= "," & vs2(3)                                         ' idx23 委託者名
            wline &= "," & miraisho.ItakusyameiCell                       ' idx23

            wline &= "," & miraisho.Iraisaki                              ' idx23 依頼先
            wline &= "," & miraisho.IraisakiCell                          ' idx23
            wline &= "," & miraisho.I_Ginkoucode                          ' idx23 銀行コード
            wline &= "," & miraisho.I_GinkoucodeCell                      ' idx23
            wline &= "," & miraisho.I_Ginkoucode1                         ' idx23 銀行コード（左から１桁目）
            wline &= "," & miraisho.I_Ginkoucode1Cell                     ' idx23
            wline &= "," & miraisho.I_Ginkoucode2                         ' idx23 銀行コード（左から２桁目）
            wline &= "," & miraisho.I_Ginkoucode2Cell                     ' idx23
            wline &= "," & miraisho.I_Ginkoucode3                         ' idx23 銀行コード（左から３桁目）
            wline &= "," & miraisho.I_Ginkoucode3Cell                     ' idx23
            wline &= "," & miraisho.I_Ginkoucode4                         ' idx23 銀行コード（左から４桁目）
            wline &= "," & miraisho.I_Ginkoucode4Cell                     ' idx23
            wline &= "," & miraisho.I_Ginkoumei                           ' idx23 銀行名
            wline &= "," & miraisho.I_GinkoumeiCell                       ' idx23
            wline &= "," & miraisho.I_Ginkou                              ' idx23 銀行名（銀行名から「銀行」などを除いた名称）
            wline &= "," & miraisho.I_GinkouCell                          ' idx23
            wline &= "," & miraisho.I_GinkouSaffix                        ' idx23 「銀行」・「信用金庫」など
            wline &= "," & miraisho.I_GinkouSaffixCell                    ' idx23
            wline &= "," & miraisho.I_Zieisinkincode                      ' idx23 自営信金コード
            wline &= "," & miraisho.I_ZieisinkincodeCell                  ' idx23
            wline &= "," & miraisho.Baitaimei                             ' idx23 媒体名
            wline &= "," & miraisho.BaitaimeiCell                         ' idx23

            gc = vs2(7).PadLeft(4, "0")
            wline &= "," & gc                                             ' idx23 銀行コード
            wline &= "," & miraisho.T_GinkoucodeCell                      ' idx23
            wline &= "," & gc(0).ToString()                               ' idx23 銀行コード（左から１桁目）
            wline &= "," & miraisho.T_Ginkoucode1Cell                     ' idx23
            wline &= "," & gc(1).ToString()                               ' idx23 銀行コード（左から２桁目）
            wline &= "," & miraisho.T_Ginkoucode2Cell                     ' idx23
            wline &= "," & gc(2).ToString()                               ' idx23 銀行コード（左から３桁目）
            wline &= "," & miraisho.T_Ginkoucode3Cell                     ' idx23
            wline &= "," & gc(3).ToString()                               ' idx23 銀行コード（左から４桁目）
            wline &= "," & miraisho.T_Ginkoucode4Cell                     ' idx23
            wline &= "," & vs(9)                                          ' idx23 銀行名
            wline &= "," & miraisho.T_GinkoumeiCell                       ' idx23
            wline &= "," & _GetGinkou(vs(9))                              ' idx23 銀行名（銀行名から「銀行」などを除いた名称）
            wline &= "," & miraisho.T_GinkouCell                          ' idx23
            wline &= "," & _GetGinkouSaffix(vs(9))                        ' idx23 「銀行」・「信用金庫」など
            wline &= "," & miraisho.T_GinkouSaffixCell                    ' idx23

            wline &= IIf(imd.HeaderFlag, "1", "0")                        ' idx21
            wline &= "," & IIf(imd.ResheetFlag, "1", "0")                 ' idx22
            wline &= "," & miraisho.BookName                              ' idx23
            wline &= "," & miraisho.ResheetType                           ' idx24
            wline &= "," & miraisho.BankNameCell                          ' idx25
            wline &= "," & miraisho.TantouCell                            ' idx26
            wline &= "," & miraisho.ItakusyacodeCell                      ' idx27
            wline &= "," & miraisho.ZieisinkincodeCell                    ' idx28
            wline &= "," & miraisho.HikiotoshikinyukikanbangouCell        ' idx29
            wline &= "," & miraisho.MeisaiTableTopRow                     ' idx30
            wline &= "," & miraisho.MeisaiTableBottomRow                  ' idx31
            wline &= "," & (miraisho.MeisaiTableTopRow + imdcount - 1)    ' idx32
            wline &= "," & miraisho.SitenmeiColumns                       ' idx33
            wline &= "," & miraisho.SitencodeColumns                      ' idx34
            wline &= "," & miraisho.KokyakucodeColumns                    ' idx35
            wline &= "," & miraisho.FunoucodeColumns                      ' idx36
            wline &= "," & miraisho.ShubetuColumns                        ' idx37
            wline &= "," & miraisho.KouzabangouColumns                    ' idx38
            wline &= "," & miraisho.KouzameigiColumns                     ' idx39
            wline &= "," & miraisho.SeikyugakuColumns                     ' idx30
            wline &= "," & miraisho.TuchoukigouColumns                    ' idx41
            wline &= "," & miraisho.TuchoubangouColumns                   ' idx42
            wline &= "," & miraisho.BikouColumns                          ' idx43

            sw.WriteLine(wline)
            wline = vbNullString
            zimd = imd
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
