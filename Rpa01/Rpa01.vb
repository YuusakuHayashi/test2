Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports System.Drawing

Public Class Rpa01 : Inherits Rpa00.RpaBase(Of Rpa01)
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
    Public SofuMeisaiColumnVLength As Double
    Public SofuMeisaiColumnWLength As Double
    Public SofuMeisaiColumnXLength As Double
    Public SofuMeisaiColumnYLength As Double
    Public SofuMeisaiColumnZLength As Double

    Public PrinterName As String
    Public PrintCreatedOnly As Boolean

    Private _BankSaffixes As List(Of String)
    Public Property BankSaffixes As List(Of String)
        Get
            If Me._BankSaffixes Is Nothing Then
                Me._BankSaffixes = New List(Of String)
            End If
            Return Me._BankSaffixes
        End Get
        Set(value As List(Of String))
            Me._BankSaffixes = value
        End Set
    End Property

    'Private _ItakusyaCodeDictionary As Dictionary(Of String, String)
    'Public Property ItakusyaCodeDictionary As Dictionary(Of String, String)
    '    Get
    '        If Me._ItakusyaCodeDictionary Is Nothing Then
    '            Me._ItakusyaCodeDictionary = New Dictionary(Of String, String)
    '            Me._ItakusyaCodeDictionary.Add("Moteki", "0029005")
    '        End If
    '        Return Me._ItakusyaCodeDictionary
    '    End Get
    '    Set(value As Dictionary(Of String, String))
    '        Me._ItakusyaCodeDictionary = value
    '    End Set
    'End Property

    Private __IraishoDirectory As String
    <JsonIgnore>
    Private ReadOnly Property _IraishoDirectory As String
        Get
            If String.IsNullOrEmpty(Me.__IraishoDirectory) Then
                Me.__IraishoDirectory = $"{Data.Project.MyRobotDirectory}\iraisho"
                If Not Directory.Exists(Me.__IraishoDirectory) Then
                    Directory.CreateDirectory(Me.__IraishoDirectory)
                End If
            End If
            Return Me.__IraishoDirectory
        End Get
    End Property

    Private __BackupDirectory As String
    Private ReadOnly Property _BackupDirectory As String
        Get
            If String.IsNullOrEmpty(Me.__BackupDirectory) Then
                Me.__BackupDirectory = $"{Data.Project.MyRobotDirectory}\backup"
                If Not Directory.Exists(Me.__BackupDirectory) Then
                    Directory.CreateDirectory(Me.__BackupDirectory)
                End If
            End If
            Return Me.__BackupDirectory
        End Get
    End Property

    Private __Work2Directory As String
    Private ReadOnly Property _Work2Directory As String
        Get
            If String.IsNullOrEmpty(Me.__Work2Directory) Then
                Me.__Work2Directory = $"{Data.Project.MyRobotDirectory}\work2"
                If Not Directory.Exists(Me.__Work2Directory) Then
                    Directory.CreateDirectory(Me.__Work2Directory)
                End If
            End If
            Return Me.__Work2Directory
        End Get
    End Property

    Private __WorkDirectory As String
    Private ReadOnly Property _WorkDirectory As String
        Get
            If String.IsNullOrEmpty(Me.__WorkDirectory) Then
                Me.__WorkDirectory = $"{Data.Project.MyRobotDirectory}\work"
                If Not Directory.Exists(Me.__WorkDirectory) Then
                    Directory.CreateDirectory(Me.__WorkDirectory)
                End If
            End If
            Return Me.__WorkDirectory
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
        Public BookName As String
        Public DistinationSheetName As String
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

    Public Overrides Function Execute(ByRef dat As Object) As Integer
        Me.RestartCount = IIf(Me.RestartCount = 0, 1, Me.RestartCount)

        Dim mutil As New Rpa00.RpaMacroUtility
        Dim putil As New Rpa00.RpaPrinterUtility
        Dim outil As New Rpa00.RpaOutlookUtility
        Dim sutil As New Rpa00.RpaShellUtility
        Dim [x] As Integer

        ' プリンター名設定
        '-----------------------------------------------------------------------------------------'
        Data.Project.UsePrinterName = Data.Project.PrinterName
        Data.Project.UsePrinterName = Data.Project.RootRobotObject.PrinterName
        Data.Project.UsePrinterName = IIf(String.IsNullOrEmpty(Me.PrinterName), Data.Project.UsePrinterName, Me.PrinterName)
        '-----------------------------------------------------------------------------------------'

        ' マスターファイルのチェック・コピー
        '-----------------------------------------------------------------------------------------'
        Dim imaster_1 = $"{Data.Project.MyRobotDirectory}\{Me.MasterCsvFileName}"
        Dim wmaster_1 = $"{Me._WorkDirectory}\{Me.MasterCsvFileName}"
        If Not Rpa00.RpaModule.FileCheckLoop(imaster_1, dat) Then
            Console.WriteLine($"中断しました")
            Console.WriteLine()
            Return 1000
        End If

        File.Copy(imaster_1, wmaster_1, True)
        '-----------------------------------------------------------------------------------------'


        ' 添付ファイルの取得・解凍・ＣＳＶデータ生成
        '-----------------------------------------------------------------------------------------'
        ' Outlook内
        Dim nameobj                                      '   添付ファイル名（オブジェクト型）
        Dim attachmentfile_1 As String = vbNullString    '   添付ファイル名（ファイル名のみ） 
        ' 添付ファイル保存後
        Dim afile_1 = vbNullString                       '   添付ファイル（フルパス）
        Dim afile_1_base = vbNullString                  '   添付ファイル（ファイル名（拡張子除く）のみ）
        Dim ixls_1 = vbNullString                        '   解凍後ファイル名
        Dim icsv_1 = $"{Me._WorkDirectory}\input.csv"    ' ＣＳＶ変換後ファイル名


        Console.WriteLine("添付ファイルを検索しています...")

        ' 添付ファイルを保存し、その名前を取得
        nameobj = outil.InvokeOutlookMacroFunction(AddressOf _GetAttachmentFile)
        If nameobj IsNot Nothing Then
            attachmentfile_1 = CType(nameobj, String)
        End If
        If String.IsNullOrEmpty(attachmentfile_1) Then
            Console.WriteLine("エラー：添付ファイルが見つかりません")
            Console.WriteLine()
            Return 1000
        End If

        For Each f In Directory.GetFiles(Me._WorkDirectory)
            If Path.GetFileName(f) = attachmentfile_1 Then
                afile_1 = f
            End If
        Next
        If String.IsNullOrEmpty(afile_1) Then
            Console.WriteLine("エラー：添付ファイルが見つかりません")
            Return 1000
        End If
        afile_1_base = Path.GetFileNameWithoutExtension(afile_1)
        ixls_1 = $"{Me._WorkDirectory}\{afile_1_base}.xls"

        ' 添付ファイルを解凍
        Console.WriteLine("添付ファイルを解凍します...")
        Call sutil.RunShell(Me.AttacheCase, $"/c {afile_1} /p={Me.PasswordOfAttacheCase} /de=1 /ow=0 /opf=0 /exit=1")
        If Not File.Exists(ixls_1) Then
            Console.WriteLine($"エラー：解凍後ファイル '{ixls_1}' がありません")
            Return 1000
        End If

        ' TEST TEST
        'ixls_1 = $"{Data.Project.MyRobotDirectory}\input.xls"

        ' ＣＳＶデータ生成
        [x] = mutil.InvokeMacro("Rpa01.CreateInputTextData", {ixls_1, icsv_1})
        '-----------------------------------------------------------------------------------------'


        ' ＣＳＶから対象得意先（モテキ）のみ抜き出し・入力ＣＳＶとのマッチング
        '-----------------------------------------------------------------------------------------'
        Dim target = "0029005"
        Dim wmaster_2 = $"{Me._WorkDirectory}\{Me.MasterCsvFileName}"
        Dim tmpcsv_1 = $"{Me._WorkDirectory}\tmp.csv"
        Dim icsv_2 = $"{Me._WorkDirectory}\input.csv"
        Dim idata1_1 = $"{Me._WorkDirectory}\idata1.csv"
        Dim idata2_1 = $"{Me._WorkDirectory}\idata2.csv"
        Call _CreateTmpCsv(wmaster_2, target, tmpcsv_1)
        Call _CompareInputCsvToTmpCsv(icsv_1, tmpcsv_1, idata1_1)
        Call _CreateIraishoMeisaiDatas(idata1_1, idata2_1)
        '-----------------------------------------------------------------------------------------'


        ' 各停止依頼書を作成
        '-----------------------------------------------------------------------------------------'
        Dim idata2_2 = $"{Me._WorkDirectory}\idata2.csv"
        Dim bookname_1 As String = vbNullString
        Dim sheetname_1 As String = vbNullString
        Dim cmp = StringComparer.Ordinal

        Console.WriteLine("停止依頼書を作成中・・・")

        ' ワークへコピー
        For Each f In Directory.GetFiles(Me._IraishoDirectory)
            File.Copy(f, $"{Me._Work2Directory}\{Path.GetFileName(f)}", True)
        Next
        [x] = mutil.InvokeMacro("Rpa01.CreateIraisho", {idata2_2})

        Me.IraishoMeisaiDatas.Sort(
            Function(before, after)
                Return (String.Compare(before.BookName, after.BookName))
            End Function
        )

        ' バックアップへコピー
        Dim creonly = Me.PrintCreatedOnly
        If Data.Transaction.Parameters.Contains("createdonly") Then
            creonly = True
        Else
            creonly = False
        End If
        If creonly Then
            For Each imd In Me.IraishoMeisaiDatas
                If bookname_1 <> imd.BookName Then
                    File.Copy(
                        imd.BookName,
                        $"{Me._BackupDirectory}\{Path.GetFileName(imd.BookName)}",
                        True
                )
                End If
            Next
        Else
            For Each f In Directory.GetFiles(Me._Work2Directory)
                File.Copy(f, $"{Me._BackupDirectory}\{Path.GetFileName(f)}", True)
            Next
        End If

        bookname_1 = vbNullString
        If Not Data.Transaction.Parameters.Contains("noprint") Then
            For Each imd In Me.IraishoMeisaiDatas
                If bookname_1 <> imd.BookName Then
                    [x] = mutil.InvokeMacro("RpaSystem.PrintOutSheet", {Data.Project.UsePrinterName, imd.BookName, imd.DistinationSheetName})
                    bookname_1 = imd.BookName
                    sheetname_1 = imd.DistinationSheetName
                End If
                If sheetname_1 <> imd.DistinationSheetName Then
                    [x] = mutil.InvokeMacro("RpaSystem.PrintOutSheet", {Data.Project.UsePrinterName, imd.BookName, imd.DistinationSheetName})
                    sheetname_1 = imd.DistinationSheetName
                End If
            Next
        End If

        Console.WriteLine("停止依頼書作成完了！")
        '-----------------------------------------------------------------------------------------'


        ' 加工済送付明細の作成
        '-----------------------------------------------------------------------------------------'
        Console.WriteLine("加工済送付明細を作成中・・・")
        Dim idx_1 = 0
        Dim wkmaster_v2 = $"{Me._WorkDirectory}\{Me.MasterCsvFileName}"
        Dim idata1_2 = $"{Me._WorkDirectory}\idata1.csv"
        Dim outxlsx_1 = $"{Me._BackupDirectory}\加工済送付明細.xlsx"
        Dim sheetname_2 = vbNullString
        Dim setting_v1(26) As Double
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnALength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnBLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnCLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnDLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnELength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnFLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnGLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnHLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnILength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnJLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnKLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnLLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnMLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnNLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnOLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnPLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnQLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnRLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnSLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnTLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnULength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnVLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnWLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnXLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnYLength : idx_1 += 1
        setting_v1(idx_1) = Data.Project.RootRobotObject.SofuMeisaiColumnZLength : idx_1 += 1

        If Me.RestartCount = 1 Then
            File.Delete(outxlsx_1)
            [x] = mutil.InvokeMacro("Rpa01.CreateSofuMeisai", {wkmaster_v2, outxlsx_1, "master", SHIFT_JIS})
        End If
        sheetname_2 = $"停止{Me.RestartCount.ToString}回目"
        [x] = mutil.InvokeMacro("Rpa01.CreateSofuMeisai", {idata1_2, outxlsx_1, sheetname_2, SHIFT_JIS, setting_v1})

        If Not Data.Transaction.Parameters.Contains("noprint") Then
            [x] = mutil.InvokeMacro("RpaSystem.PrintOutSheet", {Data.Project.UsePrinterName, outxlsx_1, sheetname_2})
        End If

        Console.WriteLine("加工済送付明細作成完了！")
        '-----------------------------------------------------------------------------------------'


        ' 件数集計ファイル作成・出力
        '-----------------------------------------------------------------------------------------'
        Dim outtxt_1 = $"{Me._BackupDirectory}\件数集計.txt"

        Console.WriteLine("件数集計ファイルを作成中・・・")
        Call _CreateSummaryFile(outtxt_1)
        If Not Data.Transaction.Parameters.Contains("noprint") Then
            Call putil.TextPrintRequest((New FileInfo(outtxt_1)), SHIFT_JIS)
        End If
        Console.WriteLine("件数集計ファイル作成完了！")
        '-----------------------------------------------------------------------------------------'


        ' 終了処理
        '-----------------------------------------------------------------------------------------'
        If (Data.Transaction.Parameters.Count > 0) And (Data.Transaction.Parameters.Contains("end")) Then
            Me.RestartCode = vbNullString
            Me.RestartCount = 1
        End If
        If (Data.Transaction.Parameters.Count > 0) And (Not Data.Transaction.Parameters.Contains("end")) Then
            Me.RestartCount += 1
        End If
        If (Data.Transaction.Parameters.Count = 0) Then
            Me.RestartCount += 1
        End If

        Call Me.Save(Data.Project.MyRobotJsonFileName, Me)
        Console.WriteLine($"{Data.Project.MyRobotJsonFileName} を更新しました")
        '-----------------------------------------------------------------------------------------'

        Console.WriteLine("処理終了！")
        Console.WriteLine()
        Return 0
    End Function


    Private Function _GetGinkouSaffix(ByVal koumei As String) As String
        ' 一度、明示的にList(Of String)に代入しないと、例外が発生する
        Dim saffixes As List(Of String) = Data.Project.RootRobotObject.BankSaffixes
        Dim saffix As String = vbNullString
        If saffixes.Count > 0 Then
            saffix = saffixes.Find(
                Function(sfx)
                    Return koumei.Contains(sfx)
                End Function
            )
        End If
        Return saffix
    End Function

    Private Function _GetGinkou(ByVal koumei As String) As String
        Dim bankname As String = koumei
        For Each saffix In Data.Project.RootRobotObject.BankSaffixes
            bankname = bankname.Replace(saffix, vbNullString)
        Next
        Return bankname
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

    ' 明細コレクションの作成およびデバッグ用ファイルの作成
    Private Sub _CreateIraishoMeisaiDatas(ByVal infile As String, ByVal otfile As String)
        Dim sr As StreamReader
        Dim sw As StreamWriter

        Dim vs1(), vs2()                                     ' ＣＳＶ各要素
        Dim line As String                                   ' 入力テキスト

        Dim imds = New List(Of IraishoMeisai)
        Dim simd As IraishoMeisai = Nothing
        Dim zimd As IraishoMeisai = Nothing

        Dim R_obj = Data.Project.RootRobotObject
        Dim R_iraishos = CType(R_obj.IraishoDatas, List(Of Iraisho))
        Dim R_iraisho As Iraisho = Nothing
        Dim R_tb As Iraisho.BankInfo = Nothing

        Dim dt As DateTime                                   ' 日付関係
        Dim culture As CultureInfo
        Dim sYYYYMMDD, sMM, sDD, wYYYYMMDD, wYY

        Dim header                                           ' ヘッダーフラグ
        Dim resheet                                          ' 改ページフラグ
        Dim row                                              ' 行位置
        Dim bookname                                         ' ブック名
        Dim srcsheetname                                     ' 作成元シート名
        Dim dstsheetname                                     ' 作成先シート名
        Dim tno, tkno, tbno                                  ' 通帳記号番号関連
        Dim zsc                                              ' 自営信金コード
        Dim gc, gm, sm                                       ' 銀行コード、銀行名、支店名

        Dim imdcount = 0                                     ' 依頼書の件数

        Dim wline As String                                  ' 出力テキスト


        ' ストリームリーダ、ライター
        sr = New StreamReader(infile, System.Text.Encoding.GetEncoding(SHIFT_JIS))
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

        ' 明細コレクションを作成
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

        ' 明細コレクションをソート
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

            ' 銀行が変わったとき依頼書を取得
            If zimd.T_Ginkoucode <> imd.T_Ginkoucode Then
                imdcount = 0
                R_iraisho = Nothing

                ' 1. 依頼書の対象銀行コードが合致する依頼書データ取得
                R_iraisho = R_iraishos.Find(
                    Function(ri)
                        Dim hit = ri.TargetBanks.Find(
                            Function(tb)
                                Return (tb.Code = imd.T_Ginkoucode)
                            End Function
                        )
                        Return IIf(hit IsNot Nothing, True, False)
                    End Function
                )

                ' 2. 対象パターン文字列群から依頼書データを取得
                If R_iraisho Is Nothing Then
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

                ' 3. デフォルト依頼書データを取得
                If R_iraisho Is Nothing Then
                    R_iraisho = R_iraishos(0)
                End If


                ' 銀行情報取得
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

            '---------------------------------------------------------------------------'
            ' ブック名を取得（印刷時に使用）
            bookname = $"{Me._Work2Directory}\{R_iraisho.BookName}"
            imd.BookName = bookname

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
            ' シート名の取得（印刷時に使用）
            imd.DistinationSheetName = dstsheetname

            ' 行数
            row = (R_iraisho.MeisaiTableTopRow - 1 + imdcount).ToString()
            '---------------------------------------------------------------------------'

            '---------------------------------------------------------------------------'
            vs2 = imd.MeisaiString.Split(",")

            ' 銀行コード、自営信金コード
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
            imd.MeisaiString = bookname                                                                                                                ' idx000 Ｅｘｃｅｌ名
            imd.MeisaiString &= "," & srcsheetname                                                                                                     ' idx001
            imd.MeisaiString &= "," & dstsheetname                                                                                                     ' idx002
            imd.MeisaiString &= "," & R_obj.Syunoukigyoucode                                                                                           ' idx003
            imd.MeisaiString &= "," & R_iraisho.SyunoukigyoucodeCell                                                                                   ' idx004
            imd.MeisaiString &= "," & R_obj.Syunoukigyoumei                                                                                            ' idx005
            imd.MeisaiString &= "," & R_iraisho.SyunoukigyoumeiCell                                                                                    ' idx006
            imd.MeisaiString &= "," & R_obj.Syunoukigyouaddress                                                                                        ' idx007
            imd.MeisaiString &= "," & R_iraisho.SyunoukigyouaddressCell                                                                                ' idx008
            imd.MeisaiString &= "," & R_obj.Syunoukigyoutel                                                                                            ' idx009
            imd.MeisaiString &= "," & R_iraisho.SyunoukigyoutelCell                                                                                    ' idx010
            imd.MeisaiString &= "," & R_obj.Syunoukigyoutantousya                                                                                      ' idx011
            imd.MeisaiString &= "," & R_iraisho.SyunoukigyoutantousyaCell                                                                              ' idx012
            imd.MeisaiString &= "," & R_obj.Haraikomisakikouzabangou                                                                                   ' idx013
            imd.MeisaiString &= "," & R_iraisho.HaraikomisakikouzabangouCell                                                                           ' idx014
            imd.MeisaiString &= "," & R_obj.Haraikomikinshubetu                                                                                        ' idx015
            imd.MeisaiString &= "," & R_iraisho.HaraikomikinshubetuCell                                                                                ' idx016
            imd.MeisaiString &= "," & sYYYYMMDD                                                                                                        ' idx017 依頼日
            imd.MeisaiString &= "," & R_iraisho.IraibiCell                                                                                             ' idx018
            imd.MeisaiString &= "," & sYYYYMMDD                                                                                                        ' idx019 作成日
            imd.MeisaiString &= "," & R_iraisho.SakuseibiCell                                                                                          ' idx020
            imd.MeisaiString &= "," & Strings.Left(sYYYYMMDD, 4)                                                                                       ' idx021 作成年（４桁）
            imd.MeisaiString &= "," & R_iraisho.SakuseibiYYYYCell                                                                                      ' idx022
            imd.MeisaiString &= "," & Strings.Mid(sYYYYMMDD, 3, 2)                                                                                     ' idx023 作成年（下２桁）
            imd.MeisaiString &= "," & R_iraisho.SakuseibiYYCell                                                                                        ' idx024
            imd.MeisaiString &= "," & wYY                                                                                                              ' idx025 作成年（和暦）
            imd.MeisaiString &= "," & R_iraisho.SakuseibiWaYYCell                                                                                      ' idx026
            imd.MeisaiString &= "," & IIf(wYY > "09", wYY, Strings.Right(wYY, 1))                                                                      ' idx027 作成年（和暦、ゼロサプレス）
            imd.MeisaiString &= "," & R_iraisho.SakuseibiWaZYCell                                                                                      ' idx028
            imd.MeisaiString &= "," & sMM                                                                                                              ' idx029 作成月
            imd.MeisaiString &= "," & R_iraisho.SakuseibiMMCell                                                                                        ' idx030
            imd.MeisaiString &= "," & IIf(sMM > "09", sMM, Strings.Right(sMM, 1))                                                                      ' idx031 作成月（ゼロサプレス）
            imd.MeisaiString &= "," & R_iraisho.SakuseibiZMCell                                                                                        ' idx032
            imd.MeisaiString &= "," & sDD                                                                                                              ' idx033 作成日
            imd.MeisaiString &= "," & R_iraisho.SakuseibiDDCell                                                                                        ' idx034
            imd.MeisaiString &= "," & IIf(sDD > "09", sDD, Strings.Right(sDD, 1))                                                                      ' idx035 作成日（ゼロサプレス）
            imd.MeisaiString &= "," & R_iraisho.SakuseibiZDCell                                                                                        ' idx036
            imd.MeisaiString &= "," & vbNullString                                                                                                     ' idx037 振替指定日
            imd.MeisaiString &= "," & R_iraisho.FurikaesiteibiCell                                                                                     ' idx038 （現在未使用）
            imd.MeisaiString &= "," & vs2(2)                                                                                                           ' idx039 委託者コード
            imd.MeisaiString &= "," & R_iraisho.ItakusyacodeCell                                                                                       ' idx040
            imd.MeisaiString &= "," & vs2(3)                                                                                                           ' idx041 委託者名
            imd.MeisaiString &= "," & R_iraisho.ItakusyameiCell                                                                                        ' idx042
            imd.MeisaiString &= "," & R_iraisho.Iraisaki                                                                                               ' idx043 依頼先
            imd.MeisaiString &= "," & R_iraisho.IraisakiCell                                                                                           ' idx044
            imd.MeisaiString &= "," & R_iraisho.I_Ginkoucode                                                                                           ' idx045 銀行コード
            imd.MeisaiString &= "," & R_iraisho.I_GinkoucodeCell                                                                                       ' idx046
            imd.MeisaiString &= "," & R_iraisho.I_Ginkoucode1                                                                                          ' idx047 銀行コード（左から１桁目）
            imd.MeisaiString &= "," & R_iraisho.I_Ginkoucode1Cell                                                                                      ' idx048
            imd.MeisaiString &= "," & R_iraisho.I_Ginkoucode2                                                                                          ' idx049 銀行コード（左から２桁目）
            imd.MeisaiString &= "," & R_iraisho.I_Ginkoucode2Cell                                                                                      ' idx050
            imd.MeisaiString &= "," & R_iraisho.I_Ginkoucode3                                                                                          ' idx051 銀行コード（左から３桁目）
            imd.MeisaiString &= "," & R_iraisho.I_Ginkoucode3Cell                                                                                      ' idx052
            imd.MeisaiString &= "," & R_iraisho.I_Ginkoucode4                                                                                          ' idx053 銀行コード（左から４桁目）
            imd.MeisaiString &= "," & R_iraisho.I_Ginkoucode4Cell                                                                                      ' idx054
            imd.MeisaiString &= "," & R_iraisho.I_Ginkoumei                                                                                            ' idx055 銀行名
            imd.MeisaiString &= "," & R_iraisho.I_GinkoumeiCell                                                                                        ' idx056
            imd.MeisaiString &= "," & R_iraisho.I_Ginkou                                                                                               ' idx057 銀行名（銀行名から「銀行」などを除いた名称）
            imd.MeisaiString &= "," & R_iraisho.I_GinkouCell                                                                                           ' idx058
            imd.MeisaiString &= "," & R_iraisho.I_GinkouSaffix                                                                                         ' idx059 「銀行」・「信用金庫」など
            imd.MeisaiString &= "," & R_iraisho.I_GinkouSaffixCell                                                                                     ' idx060
            imd.MeisaiString &= "," & R_iraisho.I_Zieisinkincode                                                                                       ' idx061 自営信金コード
            imd.MeisaiString &= "," & R_iraisho.I_ZieisinkincodeCell                                                                                   ' idx062
            imd.MeisaiString &= "," & R_iraisho.Baitaimei                                                                                              ' idx063 媒体名
            imd.MeisaiString &= "," & R_iraisho.BaitaimeiCell                                                                                          ' idx064
            imd.MeisaiString &= "," & gc                                                                                                               ' idx065 銀行コード
            imd.MeisaiString &= "," & R_iraisho.T_GinkoucodeCell                                                                                       ' idx066
            imd.MeisaiString &= "," & gc(0).ToString()                                                                                                 ' idx067 銀行コード（左から１桁目）
            imd.MeisaiString &= "," & R_iraisho.T_Ginkoucode1Cell                                                                                      ' idx068
            imd.MeisaiString &= "," & gc(1).ToString()                                                                                                 ' idx069 銀行コード（左から２桁目）
            imd.MeisaiString &= "," & R_iraisho.T_Ginkoucode2Cell                                                                                      ' idx070
            imd.MeisaiString &= "," & gc(2).ToString()                                                                                                 ' idx071 銀行コード（左から３桁目）
            imd.MeisaiString &= "," & R_iraisho.T_Ginkoucode3Cell                                                                                      ' idx072
            imd.MeisaiString &= "," & gc(3).ToString()                                                                                                 ' idx073 銀行コード（左から４桁目）
            imd.MeisaiString &= "," & R_iraisho.T_Ginkoucode4Cell                                                                                      ' idx074
            imd.MeisaiString &= "," & gm                                                                                                               ' idx075 銀行名
            imd.MeisaiString &= "," & R_iraisho.T_GinkoumeiCell                                                                                        ' idx076
            imd.MeisaiString &= "," & _GetGinkou(gm)                                                                                                   ' idx077 銀行名（銀行名から「銀行」などを除いた名称）
            imd.MeisaiString &= "," & R_iraisho.T_GinkouCell                                                                                           ' idx078
            imd.MeisaiString &= "," & _GetGinkouSaffix(gm)                                                                                             ' idx079 「銀行」・「信用金庫」など
            imd.MeisaiString &= "," & R_iraisho.T_GinkouSaffixCell                                                                                     ' idx080
            imd.MeisaiString &= "," & zsc                                                                                                              ' idx081 自営信金コード
            imd.MeisaiString &= "," & R_iraisho.T_ZieisinkincodeCell                                                                                   ' idx082
            imd.MeisaiString &= "," & vs2(4)                                                                                                           ' idx083 顧客コード
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.KokyakucodeColumn), vbNullString, R_iraisho.KokyakucodeColumn & row)          ' idx084
            imd.MeisaiString &= "," & gc                                                                                                               ' idx085 銀行コード
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.K_GinkoucodeColumn), vbNullString, R_iraisho.K_GinkoucodeColumn & row)        ' idx086
            imd.MeisaiString &= "," & vs2(8)                                                                                                           ' idx087 支店コード
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.SitencodeColumn), vbNullString, R_iraisho.SitencodeColumn & row)              ' idx088
            imd.MeisaiString &= "," & gm                                                                                                               ' idx089 銀行名
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.K_GinkoumeiColumn), vbNullString, R_iraisho.K_GinkoumeiColumn & row)          ' idx090
            imd.MeisaiString &= "," & _GetGinkou(gm)                                                                                                   ' idx091 銀行名（銀行名から「銀行」などを除いた名称）
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.K_GinkouColumn), vbNullString, R_iraisho.K_GinkouColumn & row)                ' idx092
            imd.MeisaiString &= "," & sm                                                                                                               ' idx093 支店名
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.SitenmeiColumn), vbNullString, R_iraisho.SitenmeiColumn & row)                ' idx094
            imd.MeisaiString &= "," & sm.Replace("支店", vbNullString)                                                                                 ' idx095 支店名（支店名から「支店」を除いた名称）
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.SitenColumn), vbNullString, R_iraisho.SitenColumn & row)                      ' idx096
            imd.MeisaiString &= "," & vs2(11)                                                                                                          ' idx097 預金種別
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.YokinshubetuColumn), vbNullString, R_iraisho.YokinshubetuColumn & row)        ' idx098
            imd.MeisaiString &= "," & vs2(12)                                                                                                          ' idx099 口座番号
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.KouzabangouColumn), vbNullString, R_iraisho.KouzabangouColumn & row)          ' idx100
            imd.MeisaiString &= "," & tkno                                                                                                             ' idx101 通帳記号
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.TuchoukigouColumn), vbNullString, R_iraisho.TuchoukigouColumn & row)          ' idx102
            imd.MeisaiString &= "," & tbno                                                                                                             ' idx103 通帳番号
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.TuchoubangouColumn), vbNullString, R_iraisho.TuchoubangouColumn & row)        ' idx104
            imd.MeisaiString &= "," & vs2(13)                                                                                                          ' idx105 口座名義（半角）
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.KouzameigiColumn), vbNullString, R_iraisho.KouzameigiColumn & row)            ' idx106
            imd.MeisaiString &= "," & Strings.StrConv(vs2(13), VbStrConv.Wide)                                                                         ' idx107 口座名義（全角）
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.JISKouzameigiColumn), vbNullString, R_iraisho.JISKouzameigiColumn & row)      ' idx108
            imd.MeisaiString &= "," & vs2(14)                                                                                                          ' idx109 振替金額
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.FurikaekingakuColumn), vbNullString, R_iraisho.FurikaekingakuColumn & row)    ' idx110
            imd.MeisaiString &= "," & R_iraisho.Bikou                                                                                                  ' idx111 備考
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.BikouColumn), vbNullString, R_iraisho.BikouColumn & row)                      ' idx112
            imd.MeisaiString &= "," & R_iraisho.Tekiyou                                                                                                ' idx113 適用
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.TekiyouColumn), vbNullString, R_iraisho.TekiyouColumn & row)                  ' idx114
            imd.MeisaiString &= "," & R_iraisho.Funoucode                                                                                              ' idx115 不能コード
            imd.MeisaiString &= "," & IIf(String.IsNullOrEmpty(R_iraisho.FunoucodeColumn), vbNullString, R_iraisho.FunoucodeColumn & row)              ' idx116
            imd.MeisaiString &= "," & vs2(19)                                                                                                          ' idx117 エラーコード

            sw.WriteLine(imd.MeisaiString)
            zimd = imd
        Next

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

        isr = New StreamReader(infile, System.Text.Encoding.GetEncoding(SHIFT_JIS))
        tsr = New StreamReader(tmpcsv, System.Text.Encoding.GetEncoding(SHIFT_JIS))
        sw = New StreamWriter(otfile, False, System.Text.Encoding.GetEncoding(SHIFT_JIS))

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
        sw = New StreamWriter(otfile, False, System.Text.Encoding.GetEncoding(SHIFT_JIS))

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

    Private Function _GetAttachmentFile(ByRef ns As Object) As Object
        Dim inbox, inboxfolders, sfolder, bfolder, [item]
        Dim rtn = vbNullString
        Try
            inbox = ns.GetDefaultFolder(6)
            Try
                inboxfolders = inbox.Folders
                Try
                    sfolder = inboxfolders.Item(Me.OutlookSourceInboxName)
                    bfolder = inboxfolders.Item(Me.OutlookBackupInboxName)

                    For Each i In sfolder.Items
                        If [item] Is Nothing Then
                            [item] = i
                        Else
                            If [item].CreationTime < i.CreationTime Then
                                [item] = i
                            End If
                        End If
                    Next

                    If [item] IsNot Nothing Then
                        For Each attachment In item.Attachments
                            If String.IsNullOrEmpty(Me.AttachmentFileName) Then
                                attachment.SaveAsFile($"{Me._WorkDirectory}\{attachment.FileName}")
                                rtn = attachment.FileName
                            End If
                        Next
                    End If

                    Do Until sfolder.Items.Count = 0
                        For Each i In sfolder.Items
                            Call i.Move(bfolder)
                            Exit For
                        Next
                    Loop
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                    Console.WriteLine($"受信用フォルダー、または保管用フォルダーの取得に失敗しました")
                    rtn = vbNullString
                Finally
                    Marshal.ReleaseComObject(sfolder)
                    Marshal.ReleaseComObject(bfolder)
                End Try
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Console.WriteLine($"フォルダーコレクションの取得に失敗しました")
                rtn = vbNullString
            Finally
                Marshal.ReleaseComObject(inboxfolders)
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine($"受信トレイの取得に失敗しました")
            rtn = vbNullString
        Finally
            Marshal.ReleaseComObject(inbox)
        End Try
        Return rtn
    End Function


    Public Overrides Function CanExecute(ByRef dat As Object) As Boolean
        Dim mutil As New Rpa00.RpaMacroUtility
        Dim outil As New Rpa00.RpaOutlookUtility

        If Not File.Exists(mutil.MacroFileName) Then
            Console.WriteLine($"マクロファイル '{mutil.MacroFileName}' が存在しません")
            Return False
        End If

        Dim obj As Object
        Dim ck As Boolean = True
        For Each [mod] In {"Rpa01", "RpaSystem", "RpaLibrary"}
            obj = mutil.InvokeMacroFunction("RpaSystem.IsModuleExist", {[mod]})
            If obj IsNot Nothing Then
                Dim ck2 = CType(obj, Boolean)
                If Not ck2 Then
                    Console.WriteLine($"ファイル '{mutil.MacroFileName}' 内にモジュール '{[mod]}' が存在しません")
                    ck = False
                End If
            Else
                Console.WriteLine($"モジュールのチェックに失敗しました")
                ck = False
                Exit For
            End If
        Next
        If Not ck Then
            Return False
        End If

        If String.IsNullOrEmpty(Me.AttacheCase) Then
            Console.WriteLine($"'MyRobotObject.AttacheCase' が指定されていません")
            Return False
        End If
        If Not File.Exists(Me.AttacheCase) Then
            Console.WriteLine($"アタッシュケース '{Me.AttacheCase}' が存在しません")
            Return False
        End If

        If String.IsNullOrEmpty(Me.OutlookSourceInboxName) Then
            Console.WriteLine($"'MyRobotObject.OutlookSourceInboxName' が指定されていません")
            Return False
        End If
        If String.IsNullOrEmpty(Me.OutlookBackupInboxName) Then
            Console.WriteLine($"'MyRobotObject.OutlookBackupInboxName' が指定されていません")
            Return False
        End If
        For Each [folder] In {Me.OutlookSourceInboxName, Me.OutlookBackupInboxName}
            If Not (outil.IsInboxFolderExist([folder])) Then
                Console.WriteLine($"受信トレイに '{[folder]}' が存在しません")
                Return False
            End If
        Next

        Return True
    End Function
End Class
