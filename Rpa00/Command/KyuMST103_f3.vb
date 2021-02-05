'-------------------------------
'磁気媒体届書作成　資格取得届
'-------------------------------
Public Class KyuMST103_f3
    Private priPGMTTL As String = "磁気媒体届書作成"
    Private priINPTTL As String = ""

    Private Enum RecordFormat
        JIMSH = 1
        KENPO = 2
        KOSEI = 3
    End Enum

    ' 現在未使用
    Private _Usable1ByteCodes As List(Of Byte)
    Private ReadOnly Property Usable1ByteCodes() As List(Of Byte)
        Get
            If Me._Usable1ByteCodes Is Nothing Then
                Me._Usable1ByteCodes = New List(Of Byte)
                Me._Usable1ByteCodes.Add(&HA) : Me._Usable1ByteCodes.Add(&HD)
                Me._Usable1ByteCodes.Add(&H1A)
                Me._Usable1ByteCodes.Add(&H20) : Me._Usable1ByteCodes.Add(&H2C) : Me._Usable1ByteCodes.Add(&H2D)
                Me._Usable1ByteCodes.Add(&H30) : Me._Usable1ByteCodes.Add(&H31) : Me._Usable1ByteCodes.Add(&H32) : Me._Usable1ByteCodes.Add(&H33) : Me._Usable1ByteCodes.Add(&H34) : Me._Usable1ByteCodes.Add(&H35) : Me._Usable1ByteCodes.Add(&H36) : Me._Usable1ByteCodes.Add(&H37) : Me._Usable1ByteCodes.Add(&H38) : Me._Usable1ByteCodes.Add(&H39)
                Me._Usable1ByteCodes.Add(&H41) : Me._Usable1ByteCodes.Add(&H42) : Me._Usable1ByteCodes.Add(&H43) : Me._Usable1ByteCodes.Add(&H44) : Me._Usable1ByteCodes.Add(&H45) : Me._Usable1ByteCodes.Add(&H46) : Me._Usable1ByteCodes.Add(&H47) : Me._Usable1ByteCodes.Add(&H48) : Me._Usable1ByteCodes.Add(&H49) : Me._Usable1ByteCodes.Add(&H4A) : Me._Usable1ByteCodes.Add(&H4B) : Me._Usable1ByteCodes.Add(&H4C) : Me._Usable1ByteCodes.Add(&H4D) : Me._Usable1ByteCodes.Add(&H4E) : Me._Usable1ByteCodes.Add(&H4F)
                Me._Usable1ByteCodes.Add(&H50) : Me._Usable1ByteCodes.Add(&H51) : Me._Usable1ByteCodes.Add(&H52) : Me._Usable1ByteCodes.Add(&H53) : Me._Usable1ByteCodes.Add(&H54) : Me._Usable1ByteCodes.Add(&H55) : Me._Usable1ByteCodes.Add(&H56) : Me._Usable1ByteCodes.Add(&H57) : Me._Usable1ByteCodes.Add(&H58) : Me._Usable1ByteCodes.Add(&H59) : Me._Usable1ByteCodes.Add(&H5A) : Me._Usable1ByteCodes.Add(&H5B) : Me._Usable1ByteCodes.Add(&H5D)
                Me._Usable1ByteCodes.Add(&H61) : Me._Usable1ByteCodes.Add(&H62) : Me._Usable1ByteCodes.Add(&H63) : Me._Usable1ByteCodes.Add(&H64) : Me._Usable1ByteCodes.Add(&H65) : Me._Usable1ByteCodes.Add(&H66) : Me._Usable1ByteCodes.Add(&H67) : Me._Usable1ByteCodes.Add(&H68) : Me._Usable1ByteCodes.Add(&H69) : Me._Usable1ByteCodes.Add(&H6A) : Me._Usable1ByteCodes.Add(&H6B) : Me._Usable1ByteCodes.Add(&H6C) : Me._Usable1ByteCodes.Add(&H6D) : Me._Usable1ByteCodes.Add(&H6E) : Me._Usable1ByteCodes.Add(&H6F)
                Me._Usable1ByteCodes.Add(&H70) : Me._Usable1ByteCodes.Add(&H71) : Me._Usable1ByteCodes.Add(&H72) : Me._Usable1ByteCodes.Add(&H73) : Me._Usable1ByteCodes.Add(&H74) : Me._Usable1ByteCodes.Add(&H75) : Me._Usable1ByteCodes.Add(&H76) : Me._Usable1ByteCodes.Add(&H77) : Me._Usable1ByteCodes.Add(&H78) : Me._Usable1ByteCodes.Add(&H79) : Me._Usable1ByteCodes.Add(&H7A) : Me._Usable1ByteCodes.Add(&H7A)
                Me._Usable1ByteCodes.Add(&HA6) : Me._Usable1ByteCodes.Add(&HA7) : Me._Usable1ByteCodes.Add(&HA8) : Me._Usable1ByteCodes.Add(&HA9) : Me._Usable1ByteCodes.Add(&HAA) : Me._Usable1ByteCodes.Add(&HAB) : Me._Usable1ByteCodes.Add(&HAC) : Me._Usable1ByteCodes.Add(&HAD) : Me._Usable1ByteCodes.Add(&HAE) : Me._Usable1ByteCodes.Add(&HAF)
                Me._Usable1ByteCodes.Add(&HB0) : Me._Usable1ByteCodes.Add(&HB1) : Me._Usable1ByteCodes.Add(&HB2) : Me._Usable1ByteCodes.Add(&HB3) : Me._Usable1ByteCodes.Add(&HB4) : Me._Usable1ByteCodes.Add(&HB5) : Me._Usable1ByteCodes.Add(&HB6) : Me._Usable1ByteCodes.Add(&HB7) : Me._Usable1ByteCodes.Add(&HB8) : Me._Usable1ByteCodes.Add(&HB9) : Me._Usable1ByteCodes.Add(&HBA) : Me._Usable1ByteCodes.Add(&HBB) : Me._Usable1ByteCodes.Add(&HBC) : Me._Usable1ByteCodes.Add(&HBD) : Me._Usable1ByteCodes.Add(&HBE) : Me._Usable1ByteCodes.Add(&HBF)
                Me._Usable1ByteCodes.Add(&HC0) : Me._Usable1ByteCodes.Add(&HC1) : Me._Usable1ByteCodes.Add(&HC2) : Me._Usable1ByteCodes.Add(&HC3) : Me._Usable1ByteCodes.Add(&HC4) : Me._Usable1ByteCodes.Add(&HC5) : Me._Usable1ByteCodes.Add(&HC6) : Me._Usable1ByteCodes.Add(&HC7) : Me._Usable1ByteCodes.Add(&HC8) : Me._Usable1ByteCodes.Add(&HC9) : Me._Usable1ByteCodes.Add(&HCA) : Me._Usable1ByteCodes.Add(&HCB) : Me._Usable1ByteCodes.Add(&HCC) : Me._Usable1ByteCodes.Add(&HCD) : Me._Usable1ByteCodes.Add(&HCE) : Me._Usable1ByteCodes.Add(&HCF)
                Me._Usable1ByteCodes.Add(&HD0) : Me._Usable1ByteCodes.Add(&HD1) : Me._Usable1ByteCodes.Add(&HD2) : Me._Usable1ByteCodes.Add(&HD3) : Me._Usable1ByteCodes.Add(&HD4) : Me._Usable1ByteCodes.Add(&HD5) : Me._Usable1ByteCodes.Add(&HD6) : Me._Usable1ByteCodes.Add(&HD7) : Me._Usable1ByteCodes.Add(&HD8) : Me._Usable1ByteCodes.Add(&HD9) : Me._Usable1ByteCodes.Add(&HDA) : Me._Usable1ByteCodes.Add(&HDB) : Me._Usable1ByteCodes.Add(&HDC) : Me._Usable1ByteCodes.Add(&HDD) : Me._Usable1ByteCodes.Add(&HDE) : Me._Usable1ByteCodes.Add(&HDF)
            End If
            Return Me._Usable1ByteCodes
        End Get
    End Property

    ' 現在未使用
    Private _Usable2ByteCodes As List(Of Byte())
    Private ReadOnly Property Usable2ByteCodes() As List(Of Byte())
        Get
            If Me._Usable2ByteCodes Is Nothing Then
                Me._Usable2ByteCodes = New List(Of Byte())
                Me._Usable2ByteCodes.Add(New Byte() {&H81, &H40, &H81, &H40}) ' 間隔
                Me._Usable2ByteCodes.Add(New Byte() {&H81, &H41, &H81, &H64}) ' 記述記号
                Me._Usable2ByteCodes.Add(New Byte() {&H81, &H65, &H81, &H7A}) ' 括弧記号
                Me._Usable2ByteCodes.Add(New Byte() {&H81, &H7B, &H81, &H8A}) ' 学術記号
                Me._Usable2ByteCodes.Add(New Byte() {&H81, &H8B, &H81, &H93}) ' 単位記号
                Me._Usable2ByteCodes.Add(New Byte() {&H81, &H94, &H81, &H9E}) ' 一般記号
                Me._Usable2ByteCodes.Add(New Byte() {&H81, &H9F, &H81, &HAC}) ' 一般記号２
                Me._Usable2ByteCodes.Add(New Byte() {&H82, &H4F, &H82, &H58}) ' 数字
                Me._Usable2ByteCodes.Add(New Byte() {&H82, &H60, &H82, &H79}) ' ローマ字
                Me._Usable2ByteCodes.Add(New Byte() {&H82, &H81, &H82, &H9A}) ' ローマ字２
                Me._Usable2ByteCodes.Add(New Byte() {&H82, &H9F, &H82, &HF1}) ' 平仮名
                Me._Usable2ByteCodes.Add(New Byte() {&H83, &H40, &H83, &H96}) ' 片仮名
                Me._Usable2ByteCodes.Add(New Byte() {&H83, &H9F, &H83, &HB6}) ' ギリシア文字
                Me._Usable2ByteCodes.Add(New Byte() {&H83, &HBF, &H83, &HD6}) ' ギリシア文字２
                Me._Usable2ByteCodes.Add(New Byte() {&H84, &H40, &H84, &H60}) ' ロシア文字
                Me._Usable2ByteCodes.Add(New Byte() {&H84, &H70, &H84, &H91}) ' ロシア文字２
                Me._Usable2ByteCodes.Add(New Byte() {&H88, &H9F, &H98, &H72}) ' 漢字
                Me._Usable2ByteCodes.Add(New Byte() {&H98, &H9F, &HEA, &HA4}) ' 漢字２
            End If
            Return Me._Usable2ByteCodes
        End Get
    End Property

    'Private _PreCmbNENG_F As String
    'Private _TargetNENG_F_Loaded As Boolean
    Private _TargetNENG_F As List(Of Date)
    Private ReadOnly Property TargetNENG_F() As List(Of Date)
        Get
            If Me._TargetNENG_F Is Nothing Then
                Me._TargetNENG_F = New List(Of Date)
            End If
            Return Me._TargetNENG_F
        End Get
    End Property

    'Private _PreCmbNENG_T As String
    'Private _TargetNENG_T_Loaded As Boolean
    Private _TargetNENG_T As List(Of Date)
    Private ReadOnly Property TargetNENG_T() As List(Of Date)
        Get
            If Me._TargetNENG_T Is Nothing Then
                Me._TargetNENG_T = New List(Of Date)
            End If
            Return Me._TargetNENG_T
        End Get
    End Property

    '全ての終了
    Private Sub KyuMST103_f3_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Me.WindowState = FormWindowState.Minimized '最小化
        FormUnLoad(Me, 1)
    End Sub


    'フォームロード
    Private Sub KyuMST103_f3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        RBtnFDCRE1.Checked = True
        RBtnFDCRE2.Checked = False
        RBtnFDCRE3.Checked = False

        '対象年月を配列にセット
        Read_KENNGT_COMB(sysKAISHA)

        '対象年月のコンボボックス設定
        Dim calendar As New System.Globalization.JapaneseCalendar()
        Dim culture As New System.Globalization.CultureInfo("ja-JP")
        culture.DateTimeFormat.Calendar = calendar
        CmbNENG_F.Items.Clear()
        For Each NENG In Me.TargetNENG_F
            Dim strNENG_F = NENG.ToString("ggyy年MM月dd日", culture)   '和暦表示
            CmbNENG_F.Items.Add(strNENG_F)
        Next
        CmbNENG_T.Items.Clear()
        For Each NENG In Me.TargetNENG_T
            Dim strNENG_T = NENG.ToString("ggyy年MM月dd日", culture)   '和暦表示
            CmbNENG_T.Items.Add(strNENG_T)
        Next
        CmbNENG_F.SelectedIndex = 0
        CmbNENG_T.SelectedIndex = 0

        DateTimePicker1_ValueChanged(DateTimePicker1, New EventArgs()) '和暦表示

        FormLoad(Me, 1)
    End Sub

    '' 資格取得日（ＦＲＯＭ）選択時
    'Private Sub CmbNENG_F_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbNENG_F.TextChanged
    '    If Me._TargetNENG_F_Loaded And Me._TargetNENG_T_Loaded Then
    '        If (Not String.IsNullOrEmpty(CmbNENG_F.Text)) And (Not String.IsNullOrEmpty(CmbNENG_T.Text)) Then
    '            Dim chkNENG_F = DateTime.ParseExact(CmbNENG_F.Text, "yyyy年MM月dd日", Nothing)
    '            Dim chkNENG_T = DateTime.ParseExact(CmbNENG_T.Text, "yyyy年MM月dd日", Nothing)
    '            If chkNENG_F.Date > chkNENG_T.Date Then
    '                MessageBox.Show("右の日付は左よりも後である必要があります")
    '            End If
    '        End If
    '    End If
    'End Sub

    '' 資格取得日（ＴＯ）選択時
    'Private Sub CmbNENG_T_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbNENG_T.TextChanged
    '    If Me._TargetNENG_F_Loaded And Me._TargetNENG_T_Loaded Then
    '        If (Not String.IsNullOrEmpty(CmbNENG_F.Text)) And (Not String.IsNullOrEmpty(CmbNENG_T.Text)) Then
    '            Dim chkNENG_F = DateTime.ParseExact(CmbNENG_F.Text, "yyyy年MM月dd日", Nothing)
    '            Dim chkNENG_T = DateTime.ParseExact(CmbNENG_T.Text, "yyyy年MM月dd日", Nothing)
    '            If chkNENG_F.Date > chkNENG_T.Date Then
    '                MessageBox.Show("右の日付は左よりも後である必要があります")
    '            End If
    '        End If
    '    End If
    'End Sub

    'ＦＤ作成年月日
    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim calendar As New System.Globalization.JapaneseCalendar()
        Dim culture As New System.Globalization.CultureInfo("ja-JP")
        culture.DateTimeFormat.Calendar = calendar
        sender.CustomFormat = DateTimePicker1.Value.ToString("ggyy", culture) + "年MM月dd日" '和暦表示
    End Sub

    'フォームのキーダウン
    Private Sub KyuMST103_f3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape : Call CmdF10Click() '終了
            Case Keys.F10 : Call CmdF10Click() '終了
        End Select

        Select Case e.KeyCode
            Case Keys.Down, Keys.Enter
                Me.SelectNextControl(Me.ActiveControl, True, True, True, True)  '次コントロールへ
            Case Keys.Up
                Me.SelectNextControl(Me.ActiveControl, False, True, True, True) '前コントロールへ
        End Select
    End Sub

    'フォームのキープレス
    Private Sub KyuMST103_f3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Select Case e.KeyChar
            Case Chr(Keys.Enter), Chr(Keys.Escape)
                e.Handled = True 'KeyPress イベントをキャンセルしBeep音を消音に
        End Select
    End Sub

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : CmdF10Click
    ' Function        : 終了処理(F10)
    ' Parameter       : NONE
    ' Return          : NONE
    ' Date            : 2009/09/26
    '-------------------------------------------------------------------
    Private Sub CmdF10Click()
        Me.Close()
    End Sub

    'ＦＤ作成ボタン
    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        If RBtnFDCRE1.Checked = True Then
            Call PutCSV_FDCRE1()
        End If

        If RBtnFDCRE2.Checked = True Then
            Call PutCSV_FDCRE2()
        End If

        If RBtnFDCRE3.Checked = True Then
            Call PutCSV_FDCRE3()
        End If
    End Sub

    'キャンセルボタン
    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _CheckRequired1ByteCode
    ' Function        : 要求データにおける使用可能文字のチェック（１バイト文字）、現在未使用
    ' Date            : 2020/12/15
    ' Arguments       : istr ... チェック対象文字列
    ' Return          : Boolean
    '-------------------------------------------------------------------
    Private Function _CheckRequired1ByteCode(ByVal istr As String) As Boolean
        Dim result As Boolean = True
        For Each ic In istr
            Dim b As Byte = Convert.ToByte(ic)
            If Not Me.Usable1ByteCodes.Contains(b) Then
            End If
        Next
        Return result
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _CheckRequired2ByteCode
    ' Function        : 要求データにおける使用可能文字のチェック（２バイト文字）、現在未使用
    ' Date            : 2020/12/15
    ' Arguments       : istr ... チェック対象文字列
    ' Return          : Boolean
    '-------------------------------------------------------------------
    Private Function _CheckRequired2ByteCode(ByVal istr As String) As Boolean
        Dim result As Boolean = True
        For Each ic In istr
            Dim ck As Boolean = False
            For Each range In Me.Usable2ByteCodes
                Dim fupper As Byte = range(0)
                Dim flower As Byte = range(1)
                Dim tupper As Byte = range(2)
                Dim tlower As Byte = range(3)
                Dim i2b() As Byte = System.Text.Encoding.GetEncoding("Shift-JIS").GetBytes(ic)
                If i2b(0) > tupper Then
                    Continue For
                End If
                If i2b(0) = tupper Then
                    If i2b(1) > tlower Then
                        Continue For
                    End If
                End If
                If i2b(0) < fupper Then
                    Continue For
                End If
                If i2b(0) = fupper Then
                    If i2b(1) < flower Then
                        Continue For
                    End If
                End If
                ck = True
            Next
            If Not ck Then
                result = False
                Exit For
            End If
        Next
        Return result
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f4.vb
    ' Purocedure Name : _GetAge
    ' Function        : 年齢計算
    ' Date            : 2020/12/15
    ' Arguments       : DATE_S ... 生年月日（８ケタ）
    ' Return          : Integer
    '-------------------------------------------------------------------
    Private Function _GetAge(ByVal DATE_S As String) As Integer
        Dim strTODAY As String = (DateTime.Now).ToString("yyyyMMdd")
        Dim lngTODAY As Long = Long.Parse(strTODAY)
        Dim lngDATE_S As Long = Long.Parse(DATE_S)
        Dim intAge As Integer = CType((lngTODAY - lngDATE_S) \ 10000, Integer)
        Return intAge
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _ConvertSysATTKBNIntoRequiredATTKBN
    ' Function        : システムが持つＳＩＫＡＫＵをデータレコード用の加入形態コードに変換
    ' Date            : 2020/12/14
    '-------------------------------------------------------------------
    Private Function _ConvertSysSIKAKUIntoRequiredSIKAKU(ByVal SysSIKAKU As String) As String
        Dim code As String = vbNullString
        Dim RequiredSIKAKU As String = vbNullString

        If String.IsNullOrEmpty(SysSIKAKU) Then
            Return vbNullString
        End If

        If SysSIKAKU.Length <= 1 Then
            Return "0"
        End If

        code = Strings.Left(SysSIKAKU, 2)
        If Not IsNumeric(code) Then
            Return "0"
        End If

        Select Case code
            Case "20" : RequiredSIKAKU = "01"    ' 新規
            Case "21" : RequiredSIKAKU = "04"    ' 再加入
            Case "22" : RequiredSIKAKU = "03"    ' 復活
            Case "24" : RequiredSIKAKU = "02"    ' 転入
            Case Else : RequiredSIKAKU = code
        End Select
        Return RequiredSIKAKU
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _ConvertSysATTKBNIntoRequiredATTKBN
    ' Function        : システムが持つＡＴＴＫＢＮをデータレコード用の被扶養者の有無に変換
    ' Date            : 2020/12/14
    '-------------------------------------------------------------------
    Private Function _ConvertSysATTKBNIntoRequiredATTKBN(ByVal SysATTKBN As String) As String
        Return IIf(SysATTKBN = "有１", "1", "0")
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _ConvertSysGETKBNIntoRequiredGETKBN
    ' Function        : システムが持つＧＥＴＫＢＮをデータレコード用の取得区分に変換
    ' Date            : 2020/12/14
    '-------------------------------------------------------------------
    Private Function _ConvertSysGETKBNIntoRequiredGETKBN(ByVal SysGETKBN As String) As String
        Dim RequiredGETKBN As String = vbNullString
        Select Case SysGETKBN
            Case "新１" : RequiredGETKBN = "1"                   '健康保険・厚生年金保険へ加入する者
            Case "再２" : RequiredGETKBN = "2"                   '
            Case "共３" : RequiredGETKBN = "3"                   '共済組合から公庫等へ出向する者
            Case "船４" : RequiredGETKBN = "4"                   '船員年金任意継続被保険者
            'Case "70歳以上" : RequiredGETKBN = "0"              '
            Case Else : RequiredGETKBN = "0"                     '
        End Select
        Return RequiredGETKBN
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _ConvertSysGengoIntoRequiredGengo
    ' Function        : システムが持つ元号コードをデータレコード用の元号に変換
    ' Date            : 2020/11/12
    '-------------------------------------------------------------------
    Private Function _ConvertSysGengoIntoRequiredGengo(ByVal SysGengo As String) As String
        Dim RequiredGengo As String = vbNullString
        Select Case SysGengo
            Case "1" : RequiredGengo = "1"     '明治
            Case "2" : RequiredGengo = "3"     '大正
            Case "3" : RequiredGengo = "5"     '昭和
            Case "4" : RequiredGengo = "7"     '平成
            Case "5" : RequiredGengo = "9"     '令和
            Case Else : RequiredGengo = "0"    'その他
        End Select
        _ConvertSysGengoIntoRequiredGengo = RequiredGengo
    End Function


    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _ConvertSysDateIntoRequiredDate
    ' Function        : データレコード用日付取得
    ' Date            : 2020/12/15
    ' Arguments       : [date] ... システム用日付（元号コード＋和暦）
    '                 : omit   ... 省略可能な項目か否か
    ' Return          : String(2)
    '                              1  ... レコード用元号コード
    '                              2  ... 和暦
    '-------------------------------------------------------------------
    Private Function _ConvertSysDateIntoRequiredDate(ByVal [date] As String, ByVal omit As Boolean) As String()
        Dim genDATE As String = vbNullString
        Dim nenDATE As String = vbNullString
        If [date].Length = 7 Then
            genDATE = _ConvertSysGengoIntoRequiredGengo(Strings.Left([date], 1))
            nenDATE = Mid([date], 2)
        End If
        If Not omit Then
            genDATE = IIf(String.IsNullOrEmpty(genDATE), "0", genDATE)
            nenDATE = IIf(String.IsNullOrEmpty(nenDATE), "000000", nenDATE)
        End If
        Return New String() {genDATE, nenDATE}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f4.vb
    ' Purocedure Name : _TrimRequiredLength
    ' Function        : データレコード用文字数調整（オーバーロードメソッド）
    ' Date            : 2020/12/15
    ' Arguments       : [data] ... 入力文字列
    '                 : [min]  ... 最低文字数
    '                 : [max]  ... 最高文字数
    '                 : pchar  ... パディングする文字
    '                 : flow   ... 文字数過不足の扱い
    ' Return          : String
    ' Note            : Ｘ～Ｙ文字の文字数に対応する場合
    '                   文字数が最低文字数に満たない場合(flow=True)
    '                     文字数が最低文字数を満たすよう右パディングを行う
    '                   文字数が最低文字数に満たない場合(flow=False)
    '                     文字数が最低文字数を満たすようパディング文字で埋める
    '                   文字数が最高文字数を超える場合(flow=True)
    '                     文字数を最高文字数まで左詰めする
    '                   文字数が最高文字数を超える場合(flow=False)
    '                     文字数が最低文字数を満たすようパディング文字で埋める
    '-------------------------------------------------------------------
    Private Overloads Function _TrimRequiredLength(ByVal [data] As String, _
                                                   ByVal [min] As Integer, _
                                                   ByVal [max] As Integer, _
                                                   ByVal [omit] As Boolean, _
                                                   Optional ByVal pchar As Char = " "c, _
                                                   Optional ByVal flow As Boolean = False) As String
        If String.IsNullOrEmpty([data]) Then
            If flow Then
                Return IIf(omit, vbNullString, [data].PadLeft(min, pchar))
            Else
                Return IIf(omit, vbNullString, Strings.StrDup(min, pchar))
            End If
        Else
            If [data].Length >= min And [data].Length <= max Then
                Return [data]
            Else
                If flow Then
                    If [data].Length < min Then
                        Return IIf(omit, vbNullString, [data].PadLeft(min, pchar))
                    End If
                    If [data].Length > max Then
                        Return IIf(omit, vbNullString, Strings.Left([data], max))
                    End If
                Else
                    If [data].Length < min Then
                        Return IIf(omit, vbNullString, Strings.StrDup(min, pchar))
                    End If
                    If [data].Length > max And flow = False Then
                        Return IIf(omit, vbNullString, Strings.StrDup(min, pchar))
                    End If
                End If
            End If
        End If
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f4.vb
    ' Purocedure Name : _TrimRequiredLength
    ' Function        : データレコード用文字数調整（オーバーロードメソッド）
    ' Date            : 2020/12/15
    ' Arguments       : [data] ... 入力文字列
    '                 : [len]  ... 必要文字数
    '                 : pchar  ... パディングする文字
    '                 : flow   ... 文字数過不足の扱い
    ' Return          : String
    ' Note            : Ｘ文字の文字数に対応する場合
    '                   必要文字数に満たない場合(flow=True)
    '                     文字数が必要文字数を満たすよう右パディングを行う
    '                   必要文字に満たない場合(flow=False)
    '                     文字数が必要文字数を満たすようパディング文字で埋める
    '                   必要文字数を超える場合(flow=True)
    '                     文字数を必要文字数まで左詰めする
    '                   必要文字数を超える場合(flow=False)
    '                     文字数が必要文字数を満たすようパディング文字で埋める
    '-------------------------------------------------------------------
    Private Overloads Function _TrimRequiredLength(ByVal [data] As String, _
                                                   ByVal [len] As Integer, _
                                                   ByVal [omit] As Boolean, _
                                                   Optional ByVal pchar As Char = " "c, _
                                                   Optional ByVal flow As Boolean = False) As String
        Return _TrimRequiredLength([data], [len], [len], [omit], pchar, flow)
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _GetSHIMEI
    ' Function        : 被保険者の氏名取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(2)
    '                             (0) ... カナ氏名
    '                             (1) ... 漢字氏名
    '-------------------------------------------------------------------
    Private Function _GetSHIMEI(ByVal [type] As Integer) As String()
        ' カナ名前
        Dim chkASHIME As String = Strings.Trim(sqlRdrFldGet("ASHIME", sysSQLDTRD))
        chkASHIME = Strings.StrConv(chkASHIME, VbStrConv.Narrow)

        ' 漢字名前
        Dim chkKSEI As String = Strings.Trim(sqlRdrFldGet("KSEI", sysSQLDTRD))
        Dim chkKMEI As String = Strings.Trim(sqlRdrFldGet("KMEI", sysSQLDTRD))
        Dim chkKSHIME As String = vbNullString
        chkKSHIME &= chkKSEI
        chkKSHIME &= "　"
        chkKSHIME &= chkKMEI
        chkKSHIME = Strings.StrConv(chkKSHIME, VbStrConv.Wide)
        'chkKSHIME = IIf(_CheckRequired2ByteCode(chkKSHIME), chkKSHIME, vbNullString)

        'chkASHIME = _TrimRequiredLength(chkASHIME, 1, 25, False, " "c, True)
        'chkKSHIME = _TrimRequiredLength(chkKSHIME, 1, 12, True, "　"c, False)
        Return New String() {chkASHIME, chkKSHIME}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f4.vb
    ' Purocedure Name : _GetDATE_S
    ' Function        : 生年月日の取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(2)
    '-------------------------------------------------------------------
    Private Function _GetDATE_S(ByVal [type] As Integer) As String()
        Dim chkDATE_S = ConvertDate(0, sqlRdrFldGet("DATE_S", sysSQLDTRD), 0)
        Dim tmpDATE_S = _ConvertSysDateIntoRequiredDate(chkDATE_S, False)
        Dim genDATE_S As String = tmpDATE_S(0)
        Dim nenDATE_S As String = tmpDATE_S(1)

        genDATE_S = _TrimRequiredLength(genDATE_S, 1, False, "0"c)
        nenDATE_S = _TrimRequiredLength(nenDATE_S, 6, False, "0"c)
        Return New String() {genDATE_S, nenDATE_S}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f4.vb
    ' Purocedure Name : _GetSEIBET
    ' Function        : 種別・性別の取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(1)
    ' Note            : 年金事務所提出用かつ７０歳以上被用者届のみ提出が
    '                   出が該当する場合、性別は省略する。
    '-------------------------------------------------------------------
    Private Function _GetSEIBET(ByVal [type] As Integer) As String()
        '種別（性別）
        Dim chkSEIBET As String = Strings.Trim(sqlRdrFldGet("SEIBET", sysSQLDTRD))

        '年金事務所提出用かつ７０歳以上被用者届のみ提出
        Dim chkTODOKE As String = _GetTODOKE([type])(0)
        If chkTODOKE = "1" Then
            chkSEIBET = vbNullString
        End If

        'chkSEIBET = _TrimRequiredLength(chkSEIBET, 1, True, "0"c)
        Return New String() {chkSEIBET}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f4.vb
    ' Purocedure Name : _GetGETKBN
    ' Function        : 取得区分の取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(1)
    ' Note            : 年金事務所提出用かつ７０歳以上被用者届のみ提出が
    '                   該当する場合、取得区分は省略する。
    '-------------------------------------------------------------------
    Private Function _GetGETKBN(ByVal [type] As Integer) As String()
        '取得区分
        Dim chkGETKBN As String = _ConvertSysGETKBNIntoRequiredGETKBN(sqlRdrFldGet("GETKBN", sysSQLDTRD))

        '年金事務所提出用かつ７０歳以上被用者届のみ提出
        Dim chkTODOKE As String = _GetTODOKE([type])(0)
        If chkTODOKE = "1" Then
            chkGETKBN = vbNullString
        End If

        'chkGETKBN = _TrimRequiredLength(chkGETKBN, 1, True, "0"c)
        Return New String() {chkGETKBN}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f4.vb
    ' Purocedure Name : _GetHONMYN
    ' Function        : 個人番号情報の取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(3)
    '                             (0) ... 個人番号
    '                             (1) ... 住民票住所を記載できない理由
    '                             (2) ... 備考欄
    ' Note            : 個人番号番号は省略する
    '                   厚生年金基金の場合、すべて省略する
    '                   厚生年金基金以外の場合で、住民票住所がない場合
    '                   は現状、無視する。よって現状すべて無視している
    '-------------------------------------------------------------------
    Private Function _GetHONMYN(ByVal [type] As Integer) As String()
        Dim chkHONMY1 As String = vbNullString
        Dim chkHONMY2 As String = vbNullString
        Dim chkHONMY3 As String = vbNullString
        'If [type] = RecordFormat.KOSEI Then
        'End If

        'chkHONMY1 = _TrimRequiredLength(chkHONMY1, 12, True, "0"c)
        'chkHONMY2 = _TrimRequiredLength(chkHONMY2, 1, True, "0"c)
        'chkHONMY3 = _TrimRequiredLength(chkHONMY3, 1, 37, True, "　"c)
        Return New String() {chkHONMY1, chkHONMY2, chkHONMY3}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f4.vb
    ' Purocedure Name : _GetNENKNO
    ' Function        : 基礎年金番号の取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(2)
    '                             (0) ... 課所符号
    '                             (1) ... 一連番号
    ' Note            : 健康保険組合提出の場合、省略する
    '-------------------------------------------------------------------
    Private Function _GetNENKNO(ByVal [type] As Integer) As String()
        Dim chkNENKNO As String = vbNullString
        Dim chkNENKN1 As String = vbNullString
        Dim chkNENKN2 As String = vbNullString
        If Not [type] = RecordFormat.KENPO Then
            chkNENKNO = Strings.Trim(sqlRdrFldGet("NENKNO", sysSQLDTRD))
            If chkNENKNO.Length = 10 Then
                chkNENKN1 = Strings.Mid(chkNENKNO, 1, 4) : chkNENKN2 = Strings.Mid(chkNENKNO, 5, 6)
            End If
        End If

        If [type] = RecordFormat.JIMSH Then
            chkNENKN1 = _TrimRequiredLength(chkNENKN1, 4, False, "0"c)
            chkNENKN2 = _TrimRequiredLength(chkNENKN2, 6, False, "0"c)
        Else
            chkNENKN1 = _TrimRequiredLength(chkNENKN1, 4, True, "0"c)
            chkNENKN2 = _TrimRequiredLength(chkNENKN2, 6, True, "0"c)
        End If
        Return New String() {chkNENKN1, chkNENKN2}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _GetKENNGT
    ' Function        : 資格取得日の取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(2)
    '-------------------------------------------------------------------
    Private Function _GetKENNGT(ByVal [type] As Integer) As String()
        Dim chkKENNGT = _ConvertSysDateIntoRequiredDate(ConvertDate(0, sqlRdrFldGet("KENNGT", sysSQLDTRD), 0), False)
        Dim genKENNGT = chkKENNGT(0)
        Dim nenKENNGT = chkKENNGT(1)
        ' 平成、令和以外は無効
        'genKENNGT = IIf(genKENNGT = "7" Or genKENNGT = "9", genKENNGT, " ")
        'nenKENNGT = IIf(genKENNGT = "7" Or genKENNGT = "9", nenKENNGT, "000000")

        genKENNGT = _TrimRequiredLength(genKENNGT, 1, False, "0"c)
        nenKENNGT = _TrimRequiredLength(nenKENNGT, 6, False, "0"c)
        Return New String() {genKENNGT, nenKENNGT}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _GetKINGAK
    ' Function        : 金額の取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 通貨によるものの額
    '                              2  ... 現物によるものの額
    '                              3  ... 上２つの合計
    ' Return          : String(3)
    '-------------------------------------------------------------------
    Private Function _GetKINGAK(ByVal [type] As Integer) As String()
        Dim chkKINGK1 = sqlRdrFldGet("HOSKEN", sysSQLDTRD, 1)
        chkKINGK1 = IIf(chkKINGK1 > 9999999, 9999999.ToString(), chkKINGK1.ToString())
        Dim chkKINGK2 = 0
        Dim chkKINGK3 = chkKINGK1 + chkKINGK2
        chkKINGK3 = IIf(chkKINGK3 > 9999999, 9999999.ToString(), chkKINGK3.ToString())

        'chkKINGK1 = _TrimRequiredLength(chkKINGK1, 1, 7, False, "0"c)
        'chkKINGK2 = _TrimRequiredLength(chkKINGK2, 1, 7, False, "0"c)
        'chkKINGK3 = _TrimRequiredLength(chkKINGK3, 1, 7, False, "0"c)
        Return New String() {chkKINGK1, chkKINGK2, chkKINGK3}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _GetBIKORN
    ' Function        : 備考用項目の取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(5)
    '                             (0) ... 備考欄項目１
    '                             (1) ... 備考欄項目２
    '                             (2) ... 備考欄項目３
    '                             (3) ... 備考欄項目４
    '                             (4) ... 備考欄
    '-------------------------------------------------------------------
    Private Function _GetBIKORN(ByVal [type] As Integer) As String()
        '備考欄項目１
        Dim chkDATE_S As String = sqlRdrFldGet("DATE_S", sysSQLDTRD)
        Dim chkBIKO01 As String = IIf(_GetAge(chkDATE_S) >= 70, "1", vbNullString)
        '備考欄項目２
        Dim chkBIKO02 As String = IIf(sqlRdrFldGet("MJIGYO", sysSQLDTRD) = 1, "1", vbNullString)
        '備考欄項目３
        Dim chkBIKO03 As String = IIf(sqlRdrFldGet("JITANW", sysSQLDTRD) = 1, "1", vbNullString)
        '備考欄項目４
        Dim chkBIKO04 As String = IIf(sqlRdrFldGet("SAIKYO", sysSQLDTRD) = 1, "1", vbNullString)
        '備考欄
        Dim chkBIKORN As String = Strings.Trim(sqlRdrFldGet("BIKO01", sysSQLDTRD))

        'chkBIKO01 = _TrimRequiredLength(chkBIKO01, 1, True, "0"c)
        'chkBIKO02 = _TrimRequiredLength(chkBIKO02, 1, True, "0"c)
        'chkBIKO03 = _TrimRequiredLength(chkBIKO03, 1, True, "0"c)
        'chkBIKO04 = _TrimRequiredLength(chkBIKO04, 1, True, "0"c)
        'chkBIKORN = _TrimRequiredLength(chkBIKORN, 1, 37, True, "　"c)
        Return New String() {chkBIKO01, chkBIKO02, chkBIKO03, chkBIKO04, chkBIKORN}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _GetADDRES
    ' Function        : 被保険者の住所取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(4)
    '                             (0) ... 郵便親番号
    '                             (1) ... 郵便子番号
    '                             (2) ... カナ住所
    '                             (3) ... 漢字住所
    ' Note            : 郵便番号は、個人情報の入力がない前提のため、
    '                   省略不可としている。
    '-------------------------------------------------------------------
    Private Function _GetADDRES(ByVal [type] As Integer) As String()
        ' 郵便番号
        Dim chkPOSTNO = Strings.Trim(sqlRdrFldGet("POSTNO", sysSQLDTRD))
        Dim chkPOSTN3 As String = vbNullString
        Dim chkPOSTN4 As String = vbNullString
        If chkPOSTNO.Length = 7 Then
            If IsNumeric(chkPOSTNO) Then
                chkPOSTN3 = Strings.Mid(chkPOSTNO, 1, 3)
                chkPOSTN4 = Strings.Mid(chkPOSTNO, 4)
            End If
        End If

        ' 住所（カナ）
        Dim chkAADRS1 As String = Strings.Trim(sqlRdrFldGet("GN_AADRS1", sysSQLDTRD))
        Dim chkAADRS2 As String = Strings.Trim(sqlRdrFldGet("GN_AADRS2", sysSQLDTRD))
        Dim chkAADRS3 As String = Strings.Trim(sqlRdrFldGet("GN_AADRS3", sysSQLDTRD))
        Dim chkAADRES As String = vbNullString
        chkAADRES &= chkAADRS1
        chkAADRES &= chkAADRS2
        chkAADRES &= chkAADRS3
        chkAADRES = Strings.StrConv(chkAADRES, VbStrConv.Narrow)

        ' 住所（漢字）
        Dim chkKADRS1 As String = Strings.Trim(sqlRdrFldGet("ADRS01", sysSQLDTRD))
        Dim chkKADRS2 As String = Strings.Trim(sqlRdrFldGet("ADRS02", sysSQLDTRD))
        Dim chkKADRS3 As String = Strings.Trim(sqlRdrFldGet("ADRS03", sysSQLDTRD))
        Dim chkKADRES As String = vbNullString
        chkKADRES &= chkKADRS1
        chkKADRES &= chkKADRS2
        chkKADRES &= chkKADRS3
        chkKADRES = Strings.StrConv(chkKADRES, VbStrConv.Wide)
        'chkKADRES = IIf(_CheckRequired2ByteCode(chkKADRES), chkKADRES, vbNullString)

        chkPOSTN3 = _TrimRequiredLength(chkPOSTN3, 3, False, "0"c)
        chkPOSTN4 = _TrimRequiredLength(chkPOSTN4, 4, False, "0"c)
        'chkAADRES = _TrimRequiredLength(chkAADRES, 1, 75, False, " "c, True)
        'chkKADRES = _TrimRequiredLength(chkKADRES, 1, 37, True, "　"c)
        Return New String() {chkPOSTN3, chkPOSTN4, chkAADRES, chkKADRES}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _GetTODOKE
    ' Function        : ７０歳以上被用者届のみ提出の取得
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(1)
    ' Note            : 年金事務所提出用以外の場合、省略
    '-------------------------------------------------------------------
    Private Function _GetTODOKE(ByVal [type] As Integer) As String()
        Dim chkTODOKE As String = vbNullString
        If [type] = RecordFormat.JIMSH Then
            chkTODOKE = IIf(sqlRdrFldGet("TODOKE", sysSQLDTRD) = 1, "1", vbNullString)
        End If

        'chkTODOKE = _TrimRequiredLength(chkTODOKE, 1, True, "0"c)
        Return New String() {chkTODOKE}
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : _GetBaseData
    ' Function        : 共通項目のセット
    ' Date            : 2020/12/15
    ' Arguments       : [type] ... データレコードフォーマット種類
    '                              1  ... 年金事務所提出用
    '                              2  ... 健康保険組合提出用
    '                              3  ... 厚生年金基金提出用
    ' Return          : String(33)
    '-------------------------------------------------------------------
    Private Function _GetBaseData(ByVal [type] As Integer) As String()
        '様式コード
        Dim chkYOSIKI As String = "2200700"

        '事業所整理番号
        Dim chkKUMAI2 As String = "10"
        Dim chkFDFLAG As String = vbNullString
        Dim chkFDKIGO As String = vbNullString
        Dim chkJIGYNO As String = vbNullString
        Read_SHAKAI_ITEM("FDFLAG", chkFDFLAG)
        Read_SHAKAI_ITEM("FDKIGO", chkFDKIGO)
        Read_SHAKAI_ITEM("JIGYNO", chkJIGYNO)
        'chkFDFLAG = _TrimRequiredLength(chkFDFLAG, 2, False, "0"c)
        'chkFDKIGO = _TrimRequiredLength(chkFDKIGO, 1, 4, False, " "c)
        'chkJIGYNO = _TrimRequiredLength(chkJIGYNO, 1, 5, False, "0"c)

        ' 被保険者整理番号
        Dim chkKZ3TOR As String = vbNullString

        ' 名前
        Dim tmpSHIMEI = _GetSHIMEI([type])
        Dim chkASHIME As String = tmpSHIMEI(0)
        Dim chkKSHIME As String = tmpSHIMEI(1)

        ' 生年月日
        Dim tmpDATE_S = _GetDATE_S([type])
        Dim genDATE_S As String = tmpDATE_S(0)
        Dim nenDATE_S As String = tmpDATE_S(1)

        ' 性別
        Dim tmpSEIBET = _GetSEIBET([type])
        Dim chkSEIBET As String = tmpSEIBET(0)

        ' 取得区分
        Dim tmpGETKBN = _GetGETKBN([type])
        Dim chkGETKBN As String = tmpGETKBN(0)

        ' 個人番号
        Dim tmpHONMYN = _GetHONMYN([type])
        Dim chkHONMY1 As String = tmpHONMYN(0)
        Dim chkHONMY2 As String = tmpHONMYN(1)
        Dim chkHONMY3 As String = tmpHONMYN(2)

        ' 基礎年金番号
        Dim tmpNENKNO = _GetNENKNO([type])
        Dim chkNENKN1 As String = tmpNENKNO(0)
        Dim chkNENKN2 As String = tmpNENKNO(1)

        ' 資格取得日
        Dim tmpKENNGT = _GetKENNGT([type])
        Dim genKENNGT As String = tmpKENNGT(0)
        Dim nenKENNGT As String = tmpKENNGT(1)

        ' 被扶養者の有無
        Dim chkATTKBN As String = _ConvertSysATTKBNIntoRequiredATTKBN(sqlRdrFldGet("ATTKBN", sysSQLDTRD))
        'chkATTKBN = _TrimRequiredLength(chkATTKBN, 1, False, "0"c)

        ' 通貨・現物によるものの額およびその合計
        Dim tmpKINGAK = _GetKINGAK([type])
        Dim chkKINGK1 As String = tmpKINGAK(0)
        Dim chkKINGK2 As String = tmpKINGAK(1)
        Dim chkKINGK3 As String = tmpKINGAK(2)

        ' 備考欄項目
        Dim tmpBIKORN = _GetBIKORN([type])
        Dim chkBIKO01 As String = tmpBIKORN(0)
        Dim chkBIKO02 As String = tmpBIKORN(1)
        Dim chkBIKO03 As String = tmpBIKORN(2)
        Dim chkBIKO04 As String = tmpBIKORN(3)
        Dim chkBIKORN As String = tmpBIKORN(4)

        ' 住所
        Dim tmpADDRES = _GetADDRES([type])
        Dim chkPOSTN3 As String = tmpADDRES(0)
        Dim chkPOSTN4 As String = tmpADDRES(1)
        Dim chkAADRES As String = tmpADDRES(2)
        Dim chkKADRES As String = tmpADDRES(3)

        ' ７０歳以上被用者届のみ提出
        Dim tmpTODOKE = _GetTODOKE([type])
        Dim chkTODOKE As String = tmpTODOKE(0)

        Dim valBASEDT(33 - 1) As String
        valBASEDT(1 - 1) = chkYOSIKI
        valBASEDT(2 - 1) = chkKUMAI2
        valBASEDT(3 - 1) = chkFDFLAG
        valBASEDT(4 - 1) = chkFDKIGO
        valBASEDT(5 - 1) = chkJIGYNO
        valBASEDT(6 - 1) = chkKZ3TOR
        valBASEDT(7 - 1) = chkASHIME
        valBASEDT(8 - 1) = chkKSHIME
        valBASEDT(9 - 1) = genDATE_S
        valBASEDT(10 - 1) = nenDATE_S
        valBASEDT(11 - 1) = chkSEIBET
        valBASEDT(12 - 1) = chkGETKBN
        valBASEDT(13 - 1) = chkHONMY1
        valBASEDT(14 - 1) = chkHONMY2
        valBASEDT(15 - 1) = chkHONMY3
        valBASEDT(16 - 1) = chkNENKN1
        valBASEDT(17 - 1) = chkNENKN2
        valBASEDT(18 - 1) = genKENNGT
        valBASEDT(19 - 1) = nenKENNGT
        valBASEDT(20 - 1) = chkATTKBN
        valBASEDT(21 - 1) = chkKINGK1
        valBASEDT(22 - 1) = chkKINGK2
        valBASEDT(23 - 1) = chkKINGK3
        valBASEDT(24 - 1) = chkBIKO01
        valBASEDT(25 - 1) = chkBIKO02
        valBASEDT(26 - 1) = chkBIKO03
        valBASEDT(27 - 1) = chkBIKO04
        valBASEDT(28 - 1) = chkBIKORN
        valBASEDT(29 - 1) = chkPOSTN3
        valBASEDT(30 - 1) = chkPOSTN4
        valBASEDT(31 - 1) = chkAADRES
        valBASEDT(32 - 1) = chkKADRES
        valBASEDT(33 - 1) = chkTODOKE
        Return valBASEDT
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : PutCSV_FDCRE1
    ' Function        : 社会保険事務所提出用のＣＳＶ作成
    ' Date            : 2009/11/18
    '-------------------------------------------------------------------
    Public Function PutCSV_FDCRE1() As Integer
        Dim StWrite As IO.StreamWriter
        Dim Wcnt As Integer = 0      'カウンタ
        Dim sBuf As String = ""      'ファイルに出力するデータ
        Dim strHEDKMK As String = "" 'ヘッダ名称
        Dim RtnCd As Integer = 0     '保存ダイヤログの表示チェック
        Dim strSetFLNM As String     'ダイヤログに表示するファイル場所
        Dim strPutFLNM As String     'ダイヤログで指定したファイル場所

        ' 処理年月
        Dim objITEM(0 To 1) As Object : Read_SPEC_ITEM("SHRNGT", objITEM)
        Dim priNENGET = objITEM(0)

        Dim valCOUNT1 As Integer = 0

        '対象件数の取得
        Dim calendar As New System.Globalization.JapaneseCalendar()
        Dim culture As New System.Globalization.CultureInfo("ja-JP")
        culture.DateTimeFormat.Calendar = calendar
        Dim datNENG_F As Date = DateTime.ParseExact(CmbNENG_F.Text, "ggyy年MM月dd日", culture)
        Dim datNENG_T As Date = DateTime.ParseExact(CmbNENG_T.Text, "ggyy年MM月dd日", culture)
        If datNENG_F.Date > datNENG_T.Date Then
            MSGGuide(9002, "左日付が右のものより大きくなっています" & vbCrLf & CmbNENG_F.Text & " > " & CmbNENG_T.Text, "", priPGMTTL)
            Exit Function
        End If
        Dim valNENG_F As String = datNENG_F.ToString("yyyyMMdd")
        Dim valNENG_T As String = datNENG_T.ToString("yyyyMMdd")
        Dim valCOUNT9 As Integer = GetSSKGET_COUNT(sysKAISHA, valNENG_F, valNENG_T)

        If valCOUNT9 = 0 Then
            MSGGuide(9002, "データが存在しません。" & vbCrLf & CmbNENG_F.Text, "", priPGMTTL)
            Exit Function
        End If

        PutCSV_FDCRE1 = 0

        Try
            '--------------------------------------------'
            'ファイルを保存するダイヤログボックスの設定
            '--------------------------------------------'
            ' .CSV のフォルダ指定の初期値はない
            strPutFLNM = ""
            strSetFLNM = "SHFD0006" 'ファイル名
            RtnCd = DlogSvFile("CSV", strSetFLNM, strPutFLNM) 'ダイヤログボックスの表示[mdlComm.vb]

            If RtnCd = 0 Then '０以外はキャンセルが押された
                '----------------------------------------
                'ガイダンスの表示   [DBSerchG]
                '----------------------------------------
                Call DBSerchG_Start("該当データを作成しています。", "しばらくお待ち下さい。", "")

                'CSV ﾌｧｲﾙｵｰﾌﾟﾝ上書き指定(False)
                StWrite = New IO.StreamWriter(strPutFLNM, False, System.Text.Encoding.GetEncoding("Shift-JIS"))

                'ヘッダ
                Dim valFDFLAG As String = ""
                Dim valFDKIGO As String = ""
                Dim valFDCNT As Integer = 0
                Dim valFDMCOD As String = ""
                Dim valROUMUS As String = ""
                Dim valFDJCNT As Decimal = 0
                Dim valJIGYNO As String = ""
                Dim valKUMAI As String = ""
                Dim valKIKJIG As String = ""
                Dim valNENKIK As String = ""

                Read_SHAKAI_ITEM("FDFLAG", valFDFLAG)
                Read_SHAKAI_ITEM("FDKIGO", valFDKIGO)
                Read_SHAKAI_ITEM("FDCNT", valFDCNT)
                Read_SHAKAI_ITEM("FDMCOD", valFDMCOD)
                Read_SHAKAI_ITEM("ROUMUS", valROUMUS)
                Read_SHAKAI_ITEM("FDJCNT", valFDJCNT)
                Read_SHAKAI_ITEM("JIGYNO", valJIGYNO)
                Read_SHAKAI_ITEM("KUMAI", valKUMAI)
                Read_SHAKAI_ITEM("KIKJIG", valKIKJIG)
                Read_SHAKAI_ITEM("NENKIK", valNENKIK)

                Dim valPOSTNO As String = ""
                Dim valPOST_A As String = ""
                Dim valPOST_B As String = ""
                Dim valKADRS1 As String = Read_KAISHA_ITEM(sysKAISHA, "KADRS1")
                Dim valKADRS2 As String = Read_KAISHA_ITEM(sysKAISHA, "KADRS2")
                Dim valKADRS3 As String = Read_KAISHA_ITEM(sysKAISHA, "KADRS3")
                Dim valKAINAM As String = Read_KAISHA_ITEM(sysKAISHA, "KAINAM")
                Dim valJIGNAM As String = Read_KAISHA_ITEM(sysKAISHA, "JIGNAM")
                Dim valTELPHO As String = Read_KAISHA_ITEM(sysKAISHA, "TELPHO")

                '郵便番号
                valPOSTNO = Read_KAISHA_ITEM(sysKAISHA, "POSTNO")
                valPOSTNO = Replace(valPOSTNO, "-", "")
                valPOST_A = Strings.Left(valPOSTNO, 3)
                valPOST_B = Strings.Mid(valPOSTNO, 4)

                'ＦＤ作成年月日の取得
                Dim strCREYMD As String = ""
                Dim cmpCREYMD As String = DateTime.Parse(DateTimePicker1.Text).ToString("yyyyMMdd")
                strCREYMD += DateTime.Parse(DateTimePicker1.Text).Year.ToString("0000")
                strCREYMD += DateTime.Parse(DateTimePicker1.Text).Month.ToString("00")
                strCREYMD += DateTime.Parse(DateTimePicker1.Text).Day.ToString("00")

                '事業所所在地
                Dim valKADRS9 As String = valKADRS1 & valKADRS2 & valKADRS3

                If Len(valKADRS9) > 37 Then valKADRS9 = Strings.Left(valKADRS9, 37) '事業所所在地は37文字まで
                If Len(valKAINAM) > 25 Then valKAINAM = Strings.Left(valKAINAM, 25) '事業所名称は25文字まで
                If Len(valJIGNAM) > 12 Then valJIGNAM = Strings.Left(valJIGNAM, 12) '事業主氏名は12文字まで
                If Len(valTELPHO) > 12 Then valTELPHO = Strings.Left(valTELPHO, 12) '電話番号は12文字まで

                '１行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & "10" & ","
                strHEDKMK = strHEDKMK & valFDFLAG & ","
                strHEDKMK = strHEDKMK & valFDKIGO & ","
                strHEDKMK = strHEDKMK & Strings.Format(valFDCNT, "000") & ","
                strHEDKMK = strHEDKMK & strCREYMD & ","
                strHEDKMK = strHEDKMK & valFDMCOD
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '２行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & "[kanri]"
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '３行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & valROUMUS & ","
                strHEDKMK = strHEDKMK & Strings.Format(valFDJCNT, "00")
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '４行目
                Dim WTELNO() As String
                WTELNO$ = Split(RTrim$(valTELPHO), "-")

                strHEDKMK = ""
                strHEDKMK = strHEDKMK & "10" & ","
                strHEDKMK = strHEDKMK & valFDFLAG & ","
                strHEDKMK = strHEDKMK & valFDKIGO & ","
                strHEDKMK = strHEDKMK & valJIGYNO & ","
                strHEDKMK = strHEDKMK & valPOST_A & ","
                strHEDKMK = strHEDKMK & valPOST_B & ","
                strHEDKMK = strHEDKMK & valKADRS9 & ","
                strHEDKMK = strHEDKMK & valKAINAM & ","
                strHEDKMK = strHEDKMK & valJIGNAM & ","
                strHEDKMK = strHEDKMK & WTELNO(0) & ","
                strHEDKMK = strHEDKMK & WTELNO(0) & ","
                strHEDKMK = strHEDKMK & WTELNO(0)
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '５行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & "[data]"
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '算定月変ファイルの読み込み開始---
                '算定月変ファイルの読み込み開始---
                '算定月変ファイルの読み込み開始---
                Dim strSQL As String = ""
                Dim strNAME As String = ""
                Dim strHENKOU As String = ""

                Try
                    Cursor.Current = Cursors.WaitCursor

                    strSQL = ""
                    strSQL = strSQL & "select "
                    strSQL = strSQL & "m1.SEIBET,m1.KIKINO,"
                    strSQL = strSQL & "ho.HONMYF,ho.HONMYN,"
                    strSQL = strSQL & "sg.KAISHA,sg.SHAINC,sg.SEQ,"
                    strSQL = strSQL & "sg.GETKBN,sg.ATTKBN,sg.SIKAKU,sg.BIKO01,"
                    strSQL = strSQL & "sg.KSEI,sg.KMEI,sg.ASHIME,"
                    strSQL = strSQL & "sg.DATE_S,sg.DATE_N,sg.KENNGT,sg.KENPNO,sg.KZ3TOR,"
                    strSQL = strSQL & "sg.NENKNO,sg.HOSKEN,sg.HOSNEN,sg.POSTNO,sg.ADRS01,"
                    strSQL = strSQL & "sg.ADRS02,sg.ADRS03,sg.GN_AADRS1,sg.GN_AADRS2,sg.GN_AADRS3,"
                    strSQL = strSQL & "sg.ADDDTE,sg.UPDDTE,sg.DELDTE,sg.USERNM,"
                    strSQL = strSQL & "sg.MJIGYO,sg.JITANW,sg.SAIKYO,sg.TODOKE"
                    strSQL = strSQL & " from      SSKGET_DAT as sg"
                    strSQL = strSQL & " left join SHOKIN_MS1 as m1 on sg.KAISHA=m1.KAISHA and (m1.NENGET=@NENGET) and sg.SHAINC=m1.SHAINC"
                    strSQL = strSQL & " left join TOMYDB.dbo.HONNIN_MST as ho on m1.KAISHA=ho.KAISHA and m1.SHAINC=ho.SHAINC"
                    strSQL = strSQL & " where sg.KAISHA=@KAISHA and sg.KENNGT between @NENG_F and @NENG_T"

                    sqlCommand(strSQL, 30, CommandType.Text, sysSQLCMMD)

                    'コマンドオブジェクトにトランザクションオブジェクトをセット
                    sqlCmdTran(sysSQLCMMD)

                    'パラメータのセット
                    sqlParaClr(sysSQLCMMD)
                    sqlParaCmd("@KAISHA", SqlDbType.Int, 4, sqlPDIR.Input, sysKAISHA, sysSQLPARA, sysSQLCMMD)
                    sqlParaCmd("@NENGET", SqlDbType.VarChar, 6, sqlPDIR.Input, priNENGET, sysSQLPARA, sysSQLCMMD)
                    sqlParaCmd("@NENG_F", SqlDbType.VarChar, 8, sqlPDIR.Input, valNENG_F, sysSQLPARA, sysSQLCMMD)
                    sqlParaCmd("@NENG_T", SqlDbType.VarChar, 8, sqlPDIR.Input, valNENG_T, sysSQLPARA, sysSQLCMMD)

                    sqlMSGFLAG = sqlMSG.OK
                    'sqlExecDtRdr内でロック待ち(タイムアウト時間)の状態となり
                    'タイムアウトが発生した場合は５が返る
                    adoRetn = sqlExecDtRdr(sysSQLDTRD, sysSQLCMMD)
                    Select Case adoRetn
                        Case 0
                            pubUPDTFLG = sysDATAFIND.OK

                            'ループ開始
                            While sysSQLDTRD.Read
                                valCOUNT1 += 1
                                Dim valBASEDT = _GetBaseData(RecordFormat.JIMSH)

                                '明細出力---
                                '明細出力---
                                '明細出力---
                                '20180704-------------------------------------------
                                strHEDKMK = ""
                                strHEDKMK = strHEDKMK & valBASEDT(1 - 1) & ","      '届書コード                         (項番1)
                                strHEDKMK = strHEDKMK & valBASEDT(2 - 1) & ","      '都道府県コード                     (項番2)      
                                strHEDKMK = strHEDKMK & valBASEDT(3 - 1) & ","      '郡市区符号                         (項番3)
                                strHEDKMK = strHEDKMK & valBASEDT(4 - 1) & ","      '事業所記号                         (項番4)
                                strHEDKMK = strHEDKMK & valBASEDT(5 - 1) & ","      '事業者番号                         (項番5)
                                strHEDKMK = strHEDKMK & valBASEDT(6 - 1) & ","      '被保険者整理番号                   (項番6)
                                strHEDKMK = strHEDKMK & valBASEDT(7 - 1) & ","      'カナ氏名                           (項番7)
                                strHEDKMK = strHEDKMK & valBASEDT(8 - 1) & ","      '漢字氏名                           (項番8)
                                strHEDKMK = strHEDKMK & valBASEDT(9 - 1) & ","      '生年月日の元号                     (項番9)
                                strHEDKMK = strHEDKMK & valBASEDT(10 - 1) & ","     '生年月日                           (項番10)
                                strHEDKMK = strHEDKMK & valBASEDT(11 - 1) & ","     '性別                               (項番11)
                                strHEDKMK = strHEDKMK & valBASEDT(12 - 1) & ","     '取得区分                           (項番12)
                                strHEDKMK = strHEDKMK & valBASEDT(13 - 1) & ","     '個人番号                           (項番13)
                                strHEDKMK = strHEDKMK & valBASEDT(14 - 1) & ","     '住民票住所を記載できない理由       (項番14)
                                strHEDKMK = strHEDKMK & valBASEDT(15 - 1) & ","     '備考欄                             (項番15)
                                strHEDKMK = strHEDKMK & valBASEDT(16 - 1) & ","     '基礎年金番号１                     (項番16)
                                strHEDKMK = strHEDKMK & valBASEDT(17 - 1) & ","     '基礎年金番号２                     (項番17)
                                strHEDKMK = strHEDKMK & valBASEDT(18 - 1) & ","     '資格取得年月日（元号）             (項番18) 
                                strHEDKMK = strHEDKMK & valBASEDT(19 - 1) & ","     '資格取得年月日（年月日）           (項番19)
                                strHEDKMK = strHEDKMK & valBASEDT(20 - 1) & ","     '被扶養者の有無                     (項番20)
                                strHEDKMK = strHEDKMK & valBASEDT(21 - 1) & ","     '通貨によるものの額                 (項番21)
                                strHEDKMK = strHEDKMK & valBASEDT(22 - 1) & ","     '現物によるものの額                 (項番22)
                                strHEDKMK = strHEDKMK & valBASEDT(23 - 1) & ","     '合計                               (項番23)
                                strHEDKMK = strHEDKMK & valBASEDT(24 - 1) & ","     '備考欄項目１                       (項番24)
                                strHEDKMK = strHEDKMK & valBASEDT(25 - 1) & ","     '備考欄項目２                       (項番25)
                                strHEDKMK = strHEDKMK & valBASEDT(26 - 1) & ","     '備考欄項目３                       (項番26)
                                strHEDKMK = strHEDKMK & valBASEDT(27 - 1) & ","     '備考欄項目４                       (項番27)
                                strHEDKMK = strHEDKMK & valBASEDT(28 - 1) & ","     '備考欄                             (項番28)
                                strHEDKMK = strHEDKMK & valBASEDT(29 - 1) & ","     '郵便番号（親番号）                 (項番29)
                                strHEDKMK = strHEDKMK & valBASEDT(30 - 1) & ","     '郵便番号（子番号）                 (項番30)
                                strHEDKMK = strHEDKMK & valBASEDT(31 - 1) & ","     'カナ被保険者住所                   (項番31)
                                strHEDKMK = strHEDKMK & valBASEDT(32 - 1) & ","     '住所被保険者住所                   (項番32)
                                strHEDKMK = strHEDKMK & valBASEDT(33 - 1)           '７０歳以上被用者届のみ提出         (項番33)

                                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                                Wcnt = Wcnt + 1
                                '[frmDBSerchG]
                                'pubDBSerchG_f.MinimizeBox = False
                                'pubDBSerchG_f.MaximizeBox = False
                                pubProgressMAX = valCOUNT9 'バーの表示
                                pubProgressCNT = CInt(Wcnt)
                                Call DBSerchG_Max(0, valCOUNT9)
                                Call DBSerchG_Value(CInt(Wcnt))

                                If pubCancelFlag = True Then 'キャンセルが押された
                                    MSGGuide(12, "データの抽出が", "", priPGMTTL)
                                    Exit While
                                End If
                            End While
                            'MessageBox.Show(valCOUNT1 & "件を読み込みました。")
                            'ループ終了

                        Case -1
                            pubUPDTFLG = sysDATAFIND.NG
                            MSGGuide(9002, "データが存在しません。", "", priPGMTTL)

                        Case -2 'タイムアウト-ロック待ち
                            'MSGGuide(17, "会社コード" & Format(valKAICD, "000") & "は", "", HKDPRGNM)
                            'TxtKAICD.SelectAll()
                            'TxtKAICD.Focus()
                        Case Else
                            pubUPDTFLG = sysDATAFIND.ER
                    End Select
                Catch ex As Exception
                    pubUPDTFLG = sysDATAFIND.ER
                    MSGGuide(9001, "(Message)" & ex.Message, vbCrLf & "(Source)" & ex.Source, "-ReadSHIKYU")
                Finally
                    sqlClosDtRdr(sysSQLDTRD)
                    If pubUPDTFLG = sysDATAFIND.ER Then Call Final()
                    Cursor.Current = Cursors.Default
                End Try
                '算定月変ファイルの読み込み終了---
                '算定月変ファイルの読み込み終了---
                '算定月変ファイルの読み込み終了---

                StWrite.Close() 'CSVﾌｧｲﾙのｸﾛｰｽﾞ

                Call DBSerchG_Close() '[frmDBSerchG]
                MSGGuide(25, valCOUNT1 & "件を出力しました。（ＣＳＶファイル）　", "", priPGMTTL)

            Else 'キャンセルが押された
                PutCSV_FDCRE1 = 9
            End If

        Catch ex As Exception
            PutCSV_FDCRE1 = 9
            MSGGuide(9001, "(Message)" & ex.Message, vbCrLf & "(Source)" & ex.Source, "-PutCSV_FDCRE1")
        End Try
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : PutCSV_FDCRE2
    ' Function        : 健康保険組合提出用のＣＳＶ作成
    ' Date            : 2009/11/18
    '-------------------------------------------------------------------
    Public Function PutCSV_FDCRE2() As Integer
        Dim StWrite As IO.StreamWriter
        Dim Wcnt As Integer = 0      'カウンタ
        Dim sBuf As String = ""      'ファイルに出力するデータ
        Dim strHEDKMK As String = "" 'ヘッダ名称
        Dim RtnCd As Integer = 0     '保存ダイヤログの表示チェック
        Dim strSetFLNM As String     'ダイヤログに表示するファイル場所
        Dim strPutFLNM As String     'ダイヤログで指定したファイル場所

        ' 処理年月
        Dim objITEM(0 To 1) As Object : Read_SPEC_ITEM("SHRNGT", objITEM)
        Dim priNENGET = objITEM(0)

        Dim valCOUNT1 As Integer = 0

        '対象件数の取得
        Dim calendar As New System.Globalization.JapaneseCalendar()
        Dim culture As New System.Globalization.CultureInfo("ja-JP")
        culture.DateTimeFormat.Calendar = calendar
        Dim datNENG_F As Date = DateTime.ParseExact(CmbNENG_F.Text, "ggyy年MM月dd日", culture)
        Dim datNENG_T As Date = DateTime.ParseExact(CmbNENG_T.Text, "ggyy年MM月dd日", culture)
        If datNENG_F.Date > datNENG_T.Date Then
            MSGGuide(9002, "左日付が右のものより大きくなっています" & vbCrLf & CmbNENG_F.Text & " > " & CmbNENG_T.Text, "", priPGMTTL)
            Exit Function
        End If
        Dim valNENG_F As String = datNENG_F.ToString("yyyyMMdd")
        Dim valNENG_T As String = datNENG_T.ToString("yyyyMMdd")
        Dim valCOUNT9 As Integer = GetSSKGET_COUNT(sysKAISHA, valNENG_F, valNENG_T)

        If valCOUNT9 = 0 Then
            MSGGuide(9002, "データが存在しません。" & vbCrLf & CmbNENG_F.Text, "", priPGMTTL)
            Exit Function
        End If

        PutCSV_FDCRE2 = 0

        Try
            '--------------------------------------------'
            'ファイルを保存するダイヤログボックスの設定
            '--------------------------------------------'
            ' .CSV のフォルダ指定の初期値はない
            strPutFLNM = ""
            strSetFLNM = "KPFD0006" 'ファイル名
            RtnCd = DlogSvFile("CSV", strSetFLNM, strPutFLNM) 'ダイヤログボックスの表示[mdlComm.vb]

            If RtnCd = 0 Then '０以外はキャンセルが押された
                '----------------------------------------
                'ガイダンスの表示   [DBSerchG]
                '----------------------------------------
                Call DBSerchG_Start("該当データを作成しています。", "しばらくお待ち下さい。", "")

                'CSV ﾌｧｲﾙｵｰﾌﾟﾝ上書き指定(False)
                StWrite = New IO.StreamWriter(strPutFLNM, False, System.Text.Encoding.GetEncoding("Shift-JIS"))

                'ヘッダ
                Dim valFDFLAG As String = ""
                Dim valFDKIGO As String = ""
                Dim valFDCNT As Integer = 0
                Dim valFDMCOD As String = ""
                Dim valROUMUS As String = ""
                Dim valFDJCNT As Decimal = 0
                Dim valJIGYNO As String = ""
                Dim valKUMAI As String = ""
                Dim valKIKJIG As String = ""
                Dim valNENKIK As String = ""

                Read_SHAKAI_ITEM("FDFLAG", valFDFLAG)
                Read_SHAKAI_ITEM("FDKIGO", valFDKIGO)
                Read_SHAKAI_ITEM("FDCNT", valFDCNT)
                Read_SHAKAI_ITEM("FDMCOD", valFDMCOD)
                Read_SHAKAI_ITEM("ROUMUS", valROUMUS)
                Read_SHAKAI_ITEM("FDJCNT", valFDJCNT)
                Read_SHAKAI_ITEM("JIGYNO", valJIGYNO)
                Read_SHAKAI_ITEM("KUMAI", valKUMAI)
                Read_SHAKAI_ITEM("KIKJIG", valKIKJIG)
                Read_SHAKAI_ITEM("NENKIK", valNENKIK)

                Dim valPOSTNO As String = ""
                Dim valPOST_A As String = ""
                Dim valPOST_B As String = ""
                Dim valKADRS1 As String = Read_KAISHA_ITEM(sysKAISHA, "KADRS1")
                Dim valKADRS2 As String = Read_KAISHA_ITEM(sysKAISHA, "KADRS2")
                Dim valKADRS3 As String = Read_KAISHA_ITEM(sysKAISHA, "KADRS3")
                Dim valKAINAM As String = Read_KAISHA_ITEM(sysKAISHA, "KAINAM")
                Dim valJIGNAM As String = Read_KAISHA_ITEM(sysKAISHA, "JIGNAM")
                Dim valTELPHO As String = Read_KAISHA_ITEM(sysKAISHA, "TELPHO")

                '郵便番号
                valPOSTNO = Read_KAISHA_ITEM(sysKAISHA, "POSTNO")
                valPOSTNO = Replace(valPOSTNO, "-", "")
                valPOST_A = Strings.Left(valPOSTNO, 3)
                valPOST_B = Strings.Mid(valPOSTNO, 4)

                'ＦＤ作成年月日の取得
                Dim cmpCREYMD As String = DateTime.Parse(DateTimePicker1.Text).ToString("yyyyMMdd")
                Dim strCREYMD As String = ""
                strCREYMD += DateTime.Parse(DateTimePicker1.Text).Year.ToString("0000")
                strCREYMD += DateTime.Parse(DateTimePicker1.Text).Month.ToString("00")
                strCREYMD += DateTime.Parse(DateTimePicker1.Text).Day.ToString("00")

                '事業所所在地
                Dim valKADRS9 As String = valKADRS1 & valKADRS2 & valKADRS3

                If Len(valKADRS9) > 37 Then valKADRS9 = Strings.Left(valKADRS9, 37) '事業所所在地は37文字まで
                If Len(valKAINAM) > 25 Then valKAINAM = Strings.Left(valKAINAM, 25) '事業所名称は25文字まで
                If Len(valJIGNAM) > 12 Then valJIGNAM = Strings.Left(valJIGNAM, 12) '事業主氏名は12文字まで
                If Len(valTELPHO) > 12 Then valTELPHO = Strings.Left(valTELPHO, 12) '電話番号は12文字まで

                '１行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & valKUMAI & ","
                strHEDKMK = strHEDKMK & Strings.Format(valFDCNT, "000") & ","
                strHEDKMK = strHEDKMK & strCREYMD & ","
                strHEDKMK = strHEDKMK & "" & "," '項目04
                strHEDKMK = strHEDKMK & "" & "," '項目05
                strHEDKMK = strHEDKMK & "" & "," '項目06
                strHEDKMK = strHEDKMK & "" & "," '項目07
                strHEDKMK = strHEDKMK & "" '項目08
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '２行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & "[kanri]"
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '３行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & valROUMUS & ","
                strHEDKMK = strHEDKMK & Strings.Format(valFDJCNT, "00")
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '４行目
                Dim WTELNO() As String
                WTELNO$ = Split(RTrim$(valTELPHO), "-")

                strHEDKMK = ""
                strHEDKMK = strHEDKMK & valKUMAI & ","
                strHEDKMK = strHEDKMK & valPOST_A & ","
                strHEDKMK = strHEDKMK & valPOST_B & ","
                strHEDKMK = strHEDKMK & valKADRS9 & ","
                strHEDKMK = strHEDKMK & valKAINAM & ","
                strHEDKMK = strHEDKMK & valJIGNAM & ","
                strHEDKMK = strHEDKMK & WTELNO(0) & ","
                strHEDKMK = strHEDKMK & WTELNO(0) & ","
                strHEDKMK = strHEDKMK & WTELNO(0)
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '５行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & "[data]"
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '算定月変ファイルの読み込み開始---
                '算定月変ファイルの読み込み開始---
                '算定月変ファイルの読み込み開始---
                Dim strSQL As String = ""
                Dim strNAME As String = ""
                Dim strHENKOU As String = ""

                Try
                    Cursor.Current = Cursors.WaitCursor

                    strSQL = ""
                    strSQL = strSQL & "select "
                    strSQL = strSQL & "m1.SEIBET,m1.KIKINO,"
                    strSQL = strSQL & "ho.HONMYF,ho.HONMYN,"
                    strSQL = strSQL & "sg.KAISHA,sg.SHAINC,sg.SEQ,"
                    strSQL = strSQL & "sg.GETKBN,sg.ATTKBN,sg.SIKAKU,sg.BIKO01,"
                    strSQL = strSQL & "sg.KSEI,sg.KMEI,sg.ASHIME,"
                    strSQL = strSQL & "sg.DATE_S,sg.DATE_N,sg.KENNGT,sg.KENPNO,sg.KZ3TOR,"
                    strSQL = strSQL & "sg.NENKNO,sg.HOSKEN,sg.HOSNEN,sg.POSTNO,sg.ADRS01,"
                    strSQL = strSQL & "sg.ADRS02,sg.ADRS03,sg.GN_AADRS1,sg.GN_AADRS2,sg.GN_AADRS3,"
                    strSQL = strSQL & "sg.ADDDTE,sg.UPDDTE,sg.DELDTE,sg.USERNM,"
                    strSQL = strSQL & "sg.MJIGYO,sg.JITANW,sg.SAIKYO,sg.TODOKE"
                    strSQL = strSQL & " from      SSKGET_DAT as sg"
                    strSQL = strSQL & " left join SHOKIN_MS1 as m1 on sg.KAISHA=m1.KAISHA and (m1.NENGET=@NENGET) and sg.SHAINC=m1.SHAINC"
                    strSQL = strSQL & " left join TOMYDB.dbo.HONNIN_MST as ho on m1.KAISHA=ho.KAISHA and m1.SHAINC=ho.SHAINC"
                    strSQL = strSQL & " where sg.KAISHA=@KAISHA and sg.KENNGT between @NENG_F and @NENG_T"

                    sqlCommand(strSQL, 30, CommandType.Text, sysSQLCMMD)

                    'コマンドオブジェクトにトランザクションオブジェクトをセット
                    sqlCmdTran(sysSQLCMMD)

                    'パラメータのセット
                    sqlParaClr(sysSQLCMMD)
                    sqlParaCmd("@KAISHA", SqlDbType.Int, 4, sqlPDIR.Input, sysKAISHA, sysSQLPARA, sysSQLCMMD)
                    sqlParaCmd("@NENGET", SqlDbType.VarChar, 6, sqlPDIR.Input, priNENGET, sysSQLPARA, sysSQLCMMD)
                    sqlParaCmd("@NENG_F", SqlDbType.VarChar, 8, sqlPDIR.Input, valNENG_F, sysSQLPARA, sysSQLCMMD)
                    sqlParaCmd("@NENG_T", SqlDbType.VarChar, 8, sqlPDIR.Input, valNENG_T, sysSQLPARA, sysSQLCMMD)

                    sqlMSGFLAG = sqlMSG.OK
                    'sqlExecDtRdr内でロック待ち(タイムアウト時間)の状態となり
                    'タイムアウトが発生した場合は５が返る
                    adoRetn = sqlExecDtRdr(sysSQLDTRD, sysSQLCMMD)
                    Select Case adoRetn
                        Case 0
                            pubUPDTFLG = sysDATAFIND.OK

                            'ループ開始
                            While sysSQLDTRD.Read
                                valCOUNT1 += 1
                                Dim valBASEDT = _GetBaseData(RecordFormat.KENPO)

                                '以下、健康保険組合独自項目欄

                                '事業所番号（健保組合）
                                Dim va2KOMK34 As String = valKUMAI
                                'va2KOMK34 = _TrimRequiredLength(va2KOMK34, 1, 4, False, "0"c)
                                '事業所被保険者証番号（健保組合）
                                Dim va2KOMK35 As String = Strings.Trim(sqlRdrFldGet("KENPNO", sysSQLDTRD))
                                'va2KOMK35 = _TrimRequiredLength(va2KOMK35, 1, 7, True)
                                '健保固有項目
                                Dim va2KOMK36 As String = vbNullString

                                '明細出力---
                                '明細出力---
                                '明細出力---
                                '20180704--------------------------------------------------------------------------
                                strHEDKMK = ""
                                strHEDKMK = strHEDKMK & valBASEDT(1 - 1) & ","      '届書コード                         (項番1)
                                strHEDKMK = strHEDKMK & valBASEDT(2 - 1) & ","      '都道府県コード                     (項番2)      
                                strHEDKMK = strHEDKMK & valBASEDT(3 - 1) & ","      '郡市区符号                         (項番3)
                                strHEDKMK = strHEDKMK & valBASEDT(4 - 1) & ","      '事業所記号                         (項番4)
                                strHEDKMK = strHEDKMK & valBASEDT(5 - 1) & ","      '事業者番号                         (項番5)
                                strHEDKMK = strHEDKMK & valBASEDT(6 - 1) & ","      '被保険者整理番号                   (項番6)
                                strHEDKMK = strHEDKMK & valBASEDT(7 - 1) & ","      'カナ氏名                           (項番7)
                                strHEDKMK = strHEDKMK & valBASEDT(8 - 1) & ","      '漢字氏名                           (項番8)
                                strHEDKMK = strHEDKMK & valBASEDT(9 - 1) & ","      '生年月日の元号                     (項番9)
                                strHEDKMK = strHEDKMK & valBASEDT(10 - 1) & ","     '生年月日                           (項番10)
                                strHEDKMK = strHEDKMK & valBASEDT(11 - 1) & ","     '性別                               (項番11)
                                strHEDKMK = strHEDKMK & valBASEDT(12 - 1) & ","     '取得区分                           (項番12)
                                strHEDKMK = strHEDKMK & valBASEDT(13 - 1) & ","     '個人番号                           (項番13)
                                strHEDKMK = strHEDKMK & valBASEDT(14 - 1) & ","     '住民票住所を記載できない理由       (項番14)
                                strHEDKMK = strHEDKMK & valBASEDT(15 - 1) & ","     '備考欄                             (項番15)
                                strHEDKMK = strHEDKMK & valBASEDT(16 - 1) & ","     '基礎年金番号１                     (項番16)
                                strHEDKMK = strHEDKMK & valBASEDT(17 - 1) & ","     '基礎年金番号２                     (項番17)
                                strHEDKMK = strHEDKMK & valBASEDT(18 - 1) & ","     '資格取得年月日（元号）             (項番18) 
                                strHEDKMK = strHEDKMK & valBASEDT(19 - 1) & ","     '資格取得年月日（年月日）           (項番19)
                                strHEDKMK = strHEDKMK & valBASEDT(20 - 1) & ","     '被扶養者の有無                     (項番20)
                                strHEDKMK = strHEDKMK & valBASEDT(21 - 1) & ","     '通貨によるものの額                 (項番21)
                                strHEDKMK = strHEDKMK & valBASEDT(22 - 1) & ","     '現物によるものの額                 (項番22)
                                strHEDKMK = strHEDKMK & valBASEDT(23 - 1) & ","     '合計                               (項番23)
                                strHEDKMK = strHEDKMK & valBASEDT(24 - 1) & ","     '備考欄項目１                       (項番24)
                                strHEDKMK = strHEDKMK & valBASEDT(25 - 1) & ","     '備考欄項目２                       (項番25)
                                strHEDKMK = strHEDKMK & valBASEDT(26 - 1) & ","     '備考欄項目３                       (項番26)
                                strHEDKMK = strHEDKMK & valBASEDT(27 - 1) & ","     '備考欄項目４                       (項番27)
                                strHEDKMK = strHEDKMK & valBASEDT(28 - 1) & ","     '備考欄                             (項番28)
                                strHEDKMK = strHEDKMK & valBASEDT(29 - 1) & ","     '郵便番号（親番号）                 (項番29)
                                strHEDKMK = strHEDKMK & valBASEDT(30 - 1) & ","     '郵便番号（子番号）                 (項番30)
                                strHEDKMK = strHEDKMK & valBASEDT(31 - 1) & ","     'カナ被保険者住所                   (項番31)
                                strHEDKMK = strHEDKMK & valBASEDT(32 - 1) & ","     '住所被保険者住所                   (項番32)
                                strHEDKMK = strHEDKMK & valBASEDT(33 - 1) & ","     '７０歳以上被用者届のみ提出         (項番33)
                                strHEDKMK = strHEDKMK & va2KOMK34 & ","             '事業所番号（健保組合）             (項番34)
                                strHEDKMK = strHEDKMK & va2KOMK35 & ","             '被保険者証番号（健保組合）         (項番35)
                                strHEDKMK = strHEDKMK & va2KOMK36                   '健保固有項目                       (項番36)

                                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                                Wcnt = Wcnt + 1
                                '[frmDBSerchG]
                                'pubDBSerchG_f.MinimizeBox = False
                                'pubDBSerchG_f.MaximizeBox = False
                                pubProgressMAX = valCOUNT9 'バーの表示
                                pubProgressCNT = CInt(Wcnt)
                                Call DBSerchG_Max(0, valCOUNT9)
                                Call DBSerchG_Value(CInt(Wcnt))

                                If pubCancelFlag = True Then 'キャンセルが押された
                                    MSGGuide(12, "データの抽出が", "", priPGMTTL)
                                    Exit While
                                End If
                            End While
                            'MessageBox.Show(valCOUNT1 & "件を読み込みました。")
                            'ループ終了

                        Case -1
                            pubUPDTFLG = sysDATAFIND.NG
                            MSGGuide(9002, "データが存在しません。", "", priPGMTTL)

                        Case -2 'タイムアウト-ロック待ち
                            'MSGGuide(17, "会社コード" & Format(valKAICD, "000") & "は", "", HKDPRGNM)
                            'TxtKAICD.SelectAll()
                            'TxtKAICD.Focus()
                        Case Else
                            pubUPDTFLG = sysDATAFIND.ER
                    End Select
                Catch ex As Exception
                    pubUPDTFLG = sysDATAFIND.ER
                    MSGGuide(9001, "(Message)" & ex.Message, vbCrLf & "(Source)" & ex.Source, "-ReadSHIKYU")
                Finally
                    sqlClosDtRdr(sysSQLDTRD)
                    If pubUPDTFLG = sysDATAFIND.ER Then Call Final()
                    Cursor.Current = Cursors.Default
                End Try
                '算定月変ファイルの読み込み終了---
                '算定月変ファイルの読み込み終了---
                '算定月変ファイルの読み込み終了---

                StWrite.Close() 'CSVﾌｧｲﾙのｸﾛｰｽﾞ

                Call DBSerchG_Close() '[frmDBSerchG]
                MSGGuide(25, valCOUNT1 & "件を出力しました。（ＣＳＶファイル）　", "", priPGMTTL)

            Else 'キャンセルが押された
                PutCSV_FDCRE2 = 9
            End If

        Catch ex As Exception
            PutCSV_FDCRE2 = 9
            MSGGuide(9001, "(Message)" & ex.Message, vbCrLf & "(Source)" & ex.Source, "-PutCSV_FDCRE2")
        End Try
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : PutCSV_FDCRE3
    ' Function        : 厚生年金基金提出用のＣＳＶ作成
    ' Date            : 2009/11/18
    '-------------------------------------------------------------------
    Public Function PutCSV_FDCRE3() As Integer
        Dim StWrite As IO.StreamWriter
        Dim Wcnt As Integer = 0      'カウンタ
        Dim sBuf As String = ""      'ファイルに出力するデータ
        Dim strHEDKMK As String = "" 'ヘッダ名称
        Dim RtnCd As Integer = 0     '保存ダイヤログの表示チェック
        Dim strSetFLNM As String     'ダイヤログに表示するファイル場所
        Dim strPutFLNM As String     'ダイヤログで指定したファイル場所

        ' 処理年月
        Dim objITEM(0 To 1) As Object : Read_SPEC_ITEM("SHRNGT", objITEM)
        Dim priNENGET = objITEM(0)

        Dim valCOUNT1 As Integer = 0

        '対象件数の取得
        Dim calendar As New System.Globalization.JapaneseCalendar()
        Dim culture As New System.Globalization.CultureInfo("ja-JP")
        culture.DateTimeFormat.Calendar = calendar
        Dim datNENG_F As Date = DateTime.ParseExact(CmbNENG_F.Text, "ggyy年MM月dd日", culture)
        Dim datNENG_T As Date = DateTime.ParseExact(CmbNENG_T.Text, "ggyy年MM月dd日", culture)
        If datNENG_F.Date > datNENG_T.Date Then
            MSGGuide(9002, "左日付が右のものより大きくなっています" & vbCrLf & CmbNENG_F.Text & " > " & CmbNENG_T.Text, "", priPGMTTL)
            Exit Function
        End If
        Dim valNENG_F As String = datNENG_F.ToString("yyyyMMdd")
        Dim valNENG_T As String = datNENG_T.ToString("yyyyMMdd")
        Dim valCOUNT9 As Integer = GetSSKGET_COUNT(sysKAISHA, valNENG_F, valNENG_T)

        If valCOUNT9 = 0 Then
            MSGGuide(9002, "データが存在しません。" & vbCrLf & CmbNENG_F.Text, "", priPGMTTL)
            Exit Function
        End If

        PutCSV_FDCRE3 = 0

        Try
            '--------------------------------------------'
            'ファイルを保存するダイヤログボックスの設定
            '--------------------------------------------'
            ' .CSV のフォルダ指定の初期値はない
            strPutFLNM = ""
            strSetFLNM = "KNFD0006" 'ファイル名
            RtnCd = DlogSvFile("CSV", strSetFLNM, strPutFLNM) 'ダイヤログボックスの表示[mdlComm.vb]

            If RtnCd = 0 Then '０以外はキャンセルが押された
                '----------------------------------------
                'ガイダンスの表示   [DBSerchG]
                '----------------------------------------
                Call DBSerchG_Start("該当データを作成しています。", "しばらくお待ち下さい。", "")

                'CSV ﾌｧｲﾙｵｰﾌﾟﾝ上書き指定(False)
                StWrite = New IO.StreamWriter(strPutFLNM, False, System.Text.Encoding.GetEncoding("Shift-JIS"))

                'ヘッダ
                Dim valFDFLAG As String = ""
                Dim valFDKIGO As String = ""
                Dim valFDCNT As Integer = 0
                Dim valFDMCOD As String = ""
                Dim valROUMUS As String = ""
                Dim valFDJCNT As Decimal = 0
                Dim valJIGYNO As String = ""
                Dim valKUMAI As String = ""
                Dim valKIKJIG As String = ""
                Dim valNENKIK As String = ""

                Read_SHAKAI_ITEM("FDFLAG", valFDFLAG)
                Read_SHAKAI_ITEM("FDKIGO", valFDKIGO)
                Read_SHAKAI_ITEM("FDCNT", valFDCNT)
                Read_SHAKAI_ITEM("FDMCOD", valFDMCOD)
                Read_SHAKAI_ITEM("ROUMUS", valROUMUS)
                Read_SHAKAI_ITEM("FDJCNT", valFDJCNT)
                Read_SHAKAI_ITEM("JIGYNO", valJIGYNO)
                Read_SHAKAI_ITEM("KUMAI", valKUMAI)
                Read_SHAKAI_ITEM("KIKJIG", valKIKJIG)
                Read_SHAKAI_ITEM("NENKIK", valNENKIK)

                Dim valPOSTNO As String = ""
                Dim valPOST_A As String = ""
                Dim valPOST_B As String = ""
                Dim valKADRS1 As String = Read_KAISHA_ITEM(sysKAISHA, "KADRS1")
                Dim valKADRS2 As String = Read_KAISHA_ITEM(sysKAISHA, "KADRS2")
                Dim valKADRS3 As String = Read_KAISHA_ITEM(sysKAISHA, "KADRS3")
                Dim valKAINAM As String = Read_KAISHA_ITEM(sysKAISHA, "KAINAM")
                Dim valJIGNAM As String = Read_KAISHA_ITEM(sysKAISHA, "JIGNAM")
                Dim valTELPHO As String = Read_KAISHA_ITEM(sysKAISHA, "TELPHO")

                '郵便番号
                valPOSTNO = Read_KAISHA_ITEM(sysKAISHA, "POSTNO")
                valPOSTNO = Replace(valPOSTNO, "-", "")
                valPOST_A = Strings.Left(valPOSTNO, 3)
                valPOST_B = Strings.Mid(valPOSTNO, 4)

                'ＦＤ作成年月日の取得
                Dim cmpCREYMD As String = DateTime.Parse(DateTimePicker1.Text).ToString("yyyyMMdd")
                Dim strCREYMD As String = ""
                strCREYMD += DateTime.Parse(DateTimePicker1.Text).Year.ToString("0000")
                strCREYMD += DateTime.Parse(DateTimePicker1.Text).Month.ToString("00")
                strCREYMD += DateTime.Parse(DateTimePicker1.Text).Day.ToString("00")

                '事業所所在地
                Dim valKADRS9 As String = valKADRS1 & valKADRS2 & valKADRS3

                If Len(valKADRS9) > 37 Then valKADRS9 = Strings.Left(valKADRS9, 37) '事業所所在地は37文字まで
                If Len(valKAINAM) > 25 Then valKAINAM = Strings.Left(valKAINAM, 25) '事業所名称は25文字まで
                If Len(valJIGNAM) > 12 Then valJIGNAM = Strings.Left(valJIGNAM, 12) '事業主氏名は12文字まで
                If Len(valTELPHO) > 12 Then valTELPHO = Strings.Left(valTELPHO, 12) '電話番号は12文字まで

                '１行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & valKIKJIG & ","
                strHEDKMK = strHEDKMK & Strings.Format(valFDCNT, "000") & ","
                strHEDKMK = strHEDKMK & strCREYMD & ","
                strHEDKMK = strHEDKMK & "" & "," '項目04
                strHEDKMK = strHEDKMK & "" & "," '項目05
                strHEDKMK = strHEDKMK & "" & "," '項目06
                strHEDKMK = strHEDKMK & "" & "," '項目07
                strHEDKMK = strHEDKMK & "" & "," '項目08
                strHEDKMK = strHEDKMK & "" & "," '項目09
                strHEDKMK = strHEDKMK & "" & "," '項目10
                strHEDKMK = strHEDKMK & "" & "," '項目11
                strHEDKMK = strHEDKMK & "" & "," '項目12
                strHEDKMK = strHEDKMK & "" & "," '項目13
                strHEDKMK = strHEDKMK & "" & "," '項目14
                strHEDKMK = strHEDKMK & "" & "," '項目15
                strHEDKMK = strHEDKMK & "" & "," '項目16
                strHEDKMK = strHEDKMK & "" & "," '項目17
                strHEDKMK = strHEDKMK & "" & "," '項目18
                strHEDKMK = strHEDKMK & "" & "," '項目19
                strHEDKMK = strHEDKMK & "" & "," '項目20
                strHEDKMK = strHEDKMK & "" & "," '項目21
                strHEDKMK = strHEDKMK & "" & "," '項目22
                strHEDKMK = strHEDKMK & "" & "," '項目23
                strHEDKMK = strHEDKMK & "" & "," '項目24
                strHEDKMK = strHEDKMK & "" & "," '項目25
                strHEDKMK = strHEDKMK & "" & "," '項目26
                strHEDKMK = strHEDKMK & "" & "," '項目27
                strHEDKMK = strHEDKMK & "" & "," '項目28
                strHEDKMK = strHEDKMK & "" & "," '項目29
                strHEDKMK = strHEDKMK & "" & "," '項目30
                strHEDKMK = strHEDKMK & "" & "," '項目31
                strHEDKMK = strHEDKMK & "" & "," '項目32
                strHEDKMK = strHEDKMK & "" & "," '項目33
                strHEDKMK = strHEDKMK & "" & "," '項目34
                strHEDKMK = strHEDKMK & "" & "," '項目35
                strHEDKMK = strHEDKMK & "" & "," '項目36
                strHEDKMK = strHEDKMK & "" & "," '項目37
                strHEDKMK = strHEDKMK & "" & "," '項目38
                strHEDKMK = strHEDKMK & "" & "," '項目39
                strHEDKMK = strHEDKMK & "" & "," '項目40
                strHEDKMK = strHEDKMK & "" & "," '項目41
                strHEDKMK = strHEDKMK & "" & "," '項目42
                strHEDKMK = strHEDKMK & "" & "," '項目43
                strHEDKMK = strHEDKMK & "" & "," '項目44
                strHEDKMK = strHEDKMK & "" & "," '項目45
                strHEDKMK = strHEDKMK & "" & "," '項目46
                strHEDKMK = strHEDKMK & "" & "," '項目47
                strHEDKMK = strHEDKMK & "" & "," '項目48
                strHEDKMK = strHEDKMK & "" & "," '項目49
                strHEDKMK = strHEDKMK & "" & "," '項目50
                strHEDKMK = strHEDKMK & "" & "," '項目51
                strHEDKMK = strHEDKMK & "" & "," '項目52
                strHEDKMK = strHEDKMK & "" '項目53
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '２行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & "[kanri]"
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '３行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & valROUMUS & ","
                strHEDKMK = strHEDKMK & Strings.Format(valFDJCNT, "00")
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '４行目
                Dim WTELNO() As String
                WTELNO$ = Split(RTrim$(valTELPHO), "-")



                strHEDKMK = ""
                strHEDKMK = strHEDKMK & valNENKIK & ","
                strHEDKMK = strHEDKMK & valKIKJIG & ","
                strHEDKMK = strHEDKMK & valPOST_A & ","
                strHEDKMK = strHEDKMK & valPOST_B & ","
                strHEDKMK = strHEDKMK & valKADRS9 & ","
                strHEDKMK = strHEDKMK & valKAINAM & ","
                strHEDKMK = strHEDKMK & valJIGNAM & ","
                strHEDKMK = strHEDKMK & WTELNO(0) & ","
                strHEDKMK = strHEDKMK & WTELNO(0) & ","
                strHEDKMK = strHEDKMK & WTELNO(0)
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '５行目
                strHEDKMK = ""
                strHEDKMK = strHEDKMK & "[data]"
                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                '算定月変ファイルの読み込み開始---
                '算定月変ファイルの読み込み開始---
                '算定月変ファイルの読み込み開始---
                Dim strSQL As String = ""
                Dim strNAME As String = ""
                Dim strHENKOU As String = ""

                Try
                    Cursor.Current = Cursors.WaitCursor

                    strSQL = ""
                    strSQL = strSQL & "select "
                    strSQL = strSQL & "m1.SEIBET,m1.KIKINO,"
                    strSQL = strSQL & "ho.HONMYF,ho.HONMYN,"
                    strSQL = strSQL & "sg.KAISHA,sg.SHAINC,sg.SEQ,"
                    strSQL = strSQL & "sg.GETKBN,sg.ATTKBN,sg.SIKAKU,sg.BIKO01,"
                    strSQL = strSQL & "sg.KSEI,sg.KMEI,sg.ASHIME,"
                    strSQL = strSQL & "sg.DATE_S,sg.DATE_N,sg.KENNGT,sg.KENPNO,sg.KZ3TOR,"
                    strSQL = strSQL & "sg.NENKNO,sg.HOSKEN,sg.HOSNEN,sg.POSTNO,sg.ADRS01,"
                    strSQL = strSQL & "sg.ADRS02,sg.ADRS03,sg.GN_AADRS1,sg.GN_AADRS2,sg.GN_AADRS3,"
                    strSQL = strSQL & "sg.ADDDTE,sg.UPDDTE,sg.DELDTE,sg.USERNM,"
                    strSQL = strSQL & "sg.MJIGYO,sg.JITANW,sg.SAIKYO,sg.TODOKE"
                    strSQL = strSQL & " from      SSKGET_DAT as sg"
                    strSQL = strSQL & " left join SHOKIN_MS1 as m1 on sg.KAISHA=m1.KAISHA and (m1.NENGET=@NENGET) and sg.SHAINC=m1.SHAINC"
                    strSQL = strSQL & " left join TOMYDB.dbo.HONNIN_MST as ho on m1.KAISHA=ho.KAISHA and m1.SHAINC=ho.SHAINC"
                    strSQL = strSQL & " where sg.KAISHA=@KAISHA and sg.KENNGT between @NENG_F and @NENG_T"

                    sqlCommand(strSQL, 30, CommandType.Text, sysSQLCMMD)

                    'コマンドオブジェクトにトランザクションオブジェクトをセット
                    sqlCmdTran(sysSQLCMMD)

                    'パラメータのセット
                    sqlParaClr(sysSQLCMMD)
                    sqlParaCmd("@KAISHA", SqlDbType.Int, 4, sqlPDIR.Input, sysKAISHA, sysSQLPARA, sysSQLCMMD)
                    sqlParaCmd("@NENGET", SqlDbType.VarChar, 6, sqlPDIR.Input, priNENGET, sysSQLPARA, sysSQLCMMD)
                    sqlParaCmd("@NENG_F", SqlDbType.VarChar, 8, sqlPDIR.Input, valNENG_F, sysSQLPARA, sysSQLCMMD)
                    sqlParaCmd("@NENG_T", SqlDbType.VarChar, 8, sqlPDIR.Input, valNENG_T, sysSQLPARA, sysSQLCMMD)

                    sqlMSGFLAG = sqlMSG.OK
                    'sqlExecDtRdr内でロック待ち(タイムアウト時間)の状態となり
                    'タイムアウトが発生した場合は５が返る
                    adoRetn = sqlExecDtRdr(sysSQLDTRD, sysSQLCMMD)
                    Select Case adoRetn
                        Case 0
                            pubUPDTFLG = sysDATAFIND.OK

                            'ループ開始
                            While sysSQLDTRD.Read
                                valCOUNT1 += 1
                                Dim valBASEDT = _GetBaseData(RecordFormat.KOSEI)

                                '基金加入員番号

                                '基金番号
                                Dim va2KOMK34 As String = valNENKIK
                                'va2KOMK34 = _TrimRequiredLength(va2KOMK34, 1, 4, False, "0"c)
                                '事業所番号
                                Dim va2KOMK35 As String = valKIKJIG
                                'va2KOMK35 = _TrimRequiredLength(va2KOMK35, 1, 7, False)
                                '加入員番号
                                Dim va2KOMK36 As String = Strings.Trim(sqlRdrFldGet("KIKINO", sysSQLDTRD))
                                'va2KOMK36 = _TrimRequiredLength(va2KOMK36, 1, 11, True)
                                '加入形態（所得事由）コード
                                Dim va2KOMK37 As String = _ConvertSysSIKAKUIntoRequiredSIKAKU(sqlRdrFldGet("SIKAKU", sysSQLDTRD))
                                'If chkKOMK37.Length <= 1 Then
                                '    chkKOMK37 = vbNullString
                                'Else
                                '    chkKOMK37 = Strings.Left(chkKOMK37, 2)
                                '    chkKOMK37 = IIf(IsNumeric(chkKOMK37), chkKOMK37, vbNullString)
                                'End If
                                'If String.IsNullOrEmpty(chkKOMK37) Then
                                '    va2KOMK37 = vbNullString
                                'Else
                                '    Select Case chkKOMK37
                                '        Case "20" : va2KOMK37 = "01"         ' 新規
                                '        Case "21" : va2KOMK37 = "04"         ' 再加入
                                '        Case "22" : va2KOMK37 = "03"         ' 復活
                                '        Case "24" : va2KOMK37 = "02"         ' 転入
                                '        Case Else : va2KOMK37 = chkKOMK37    ' その他
                                '    End Select
                                'End If
                                'va2KOMK37 = _TrimRequiredLength(va2KOMK37, 1, 2, False, "0"c)
                                '入社年月日（元号、年月日）
                                Dim chkDATE_N = _ConvertSysDateIntoRequiredDate(ConvertDate(0, sqlRdrFldGet("DATE_N", sysSQLDTRD), 0), False)
                                Dim va2KOMK38 = chkDATE_N(0)
                                Dim va2KOMK39 = chkDATE_N(1)
                                va2KOMK38 = _TrimRequiredLength(va2KOMK38, 1, False, "0"c)
                                va2KOMK39 = _TrimRequiredLength(va2KOMK39, 6, False, "0"c)
                                '加算適用の有無
                                Dim va2KOMK40 As String = vbNullString
                                '適用形態（取得事由）コード
                                Dim va2KOMK41 As String = vbNullString
                                '加算給与月額
                                Dim va2KOMK42 As String = vbNullString
                                '標準給与月額
                                Dim va2KOMK43 As String = vbNullString
                                Dim chkHOSNEN = sqlRdrFldGet("HOSNEN", sysSQLDTRD, 1) \ 1000
                                va2KOMK43 = IIf(chkHOSNEN >= 10000, "9999", chkHOSNEN.ToString())
                                'va2KOMK43 = _TrimRequiredLength(va2KOMK43, 1, 4, True, "0"c)
                                '第２加算標準月額
                                Dim va2KOMK44 As String = vbNullString
                                '第２加算標準給与月額
                                Dim va2KOMK45 As String = vbNullString
                                '基金固有項目１
                                Dim va2KOMK46 As String = vbNullString
                                '基金固有項目２
                                Dim va2KOMK47 As String = vbNullString
                                '基金固有項目３
                                Dim va2KOMK48 As String = vbNullString
                                '基金固有項目４
                                Dim va2KOMK49 As String = vbNullString
                                '基金固有項目５
                                Dim va2KOMK50 As String = vbNullString
                                '基金固有項目６
                                Dim va2KOMK51 As String = vbNullString
                                '基金固有項目７
                                Dim va2KOMK52 As String = vbNullString
                                '基金固有項目８
                                Dim va2KOMK53 As String = vbNullString
                                '基金固有項目９
                                Dim va2KOMK54 As String = vbNullString
                                '基金固有項目１０
                                Dim va2KOMK55 As String = vbNullString


                                '明細出力---
                                '明細出力---
                                '明細出力---
                                '20180704--------------------------------------------------------------------------
                                strHEDKMK = ""
                                strHEDKMK = strHEDKMK & valBASEDT(1 - 1) & ","      '届書コード                         (項番1)
                                strHEDKMK = strHEDKMK & valBASEDT(2 - 1) & ","      '都道府県コード                     (項番2)      
                                strHEDKMK = strHEDKMK & valBASEDT(3 - 1) & ","      '郡市区符号                         (項番3)
                                strHEDKMK = strHEDKMK & valBASEDT(4 - 1) & ","      '事業所記号                         (項番4)
                                strHEDKMK = strHEDKMK & valBASEDT(5 - 1) & ","      '事業者番号                         (項番5)
                                strHEDKMK = strHEDKMK & valBASEDT(6 - 1) & ","      '被保険者整理番号                   (項番6)
                                strHEDKMK = strHEDKMK & valBASEDT(7 - 1) & ","      'カナ氏名                           (項番7)
                                strHEDKMK = strHEDKMK & valBASEDT(8 - 1) & ","      '漢字氏名                           (項番8)
                                strHEDKMK = strHEDKMK & valBASEDT(9 - 1) & ","      '生年月日の元号                     (項番9)
                                strHEDKMK = strHEDKMK & valBASEDT(10 - 1) & ","     '生年月日                           (項番10)
                                strHEDKMK = strHEDKMK & valBASEDT(11 - 1) & ","     '性別                               (項番11)
                                strHEDKMK = strHEDKMK & valBASEDT(12 - 1) & ","     '取得区分                           (項番12)
                                strHEDKMK = strHEDKMK & valBASEDT(13 - 1) & ","     '個人番号                           (項番13)
                                strHEDKMK = strHEDKMK & valBASEDT(14 - 1) & ","     '住民票住所を記載できない理由       (項番14)
                                strHEDKMK = strHEDKMK & valBASEDT(15 - 1) & ","     '備考欄                             (項番15)
                                strHEDKMK = strHEDKMK & valBASEDT(16 - 1) & ","     '基礎年金番号１                     (項番16)
                                strHEDKMK = strHEDKMK & valBASEDT(17 - 1) & ","     '基礎年金番号２                     (項番17)
                                strHEDKMK = strHEDKMK & valBASEDT(18 - 1) & ","     '資格取得年月日（元号）             (項番18) 
                                strHEDKMK = strHEDKMK & valBASEDT(19 - 1) & ","     '資格取得年月日（年月日）           (項番19)
                                strHEDKMK = strHEDKMK & valBASEDT(20 - 1) & ","     '被扶養者の有無                     (項番20)
                                strHEDKMK = strHEDKMK & valBASEDT(21 - 1) & ","     '通貨によるものの額                 (項番21)
                                strHEDKMK = strHEDKMK & valBASEDT(22 - 1) & ","     '現物によるものの額                 (項番22)
                                strHEDKMK = strHEDKMK & valBASEDT(23 - 1) & ","     '合計                               (項番23)
                                strHEDKMK = strHEDKMK & valBASEDT(24 - 1) & ","     '備考欄項目１                       (項番24)
                                strHEDKMK = strHEDKMK & valBASEDT(25 - 1) & ","     '備考欄項目２                       (項番25)
                                strHEDKMK = strHEDKMK & valBASEDT(26 - 1) & ","     '備考欄項目３                       (項番26)
                                strHEDKMK = strHEDKMK & valBASEDT(27 - 1) & ","     '備考欄項目４                       (項番27)
                                strHEDKMK = strHEDKMK & valBASEDT(28 - 1) & ","     '備考欄                             (項番28)
                                strHEDKMK = strHEDKMK & valBASEDT(29 - 1) & ","     '郵便番号（親番号）                 (項番29)
                                strHEDKMK = strHEDKMK & valBASEDT(30 - 1) & ","     '郵便番号（子番号）                 (項番30)
                                strHEDKMK = strHEDKMK & valBASEDT(31 - 1) & ","     'カナ被保険者住所                   (項番31)
                                strHEDKMK = strHEDKMK & valBASEDT(32 - 1) & ","     '住所被保険者住所                   (項番32)
                                strHEDKMK = strHEDKMK & valBASEDT(33 - 1) & ","     '７０歳以上被用者届のみ提出         (項番33)
                                strHEDKMK = strHEDKMK & va2KOMK34 & ","             '基金番号                           (項番34)
                                strHEDKMK = strHEDKMK & va2KOMK35 & ","             '事業所番号                         (項番35)
                                strHEDKMK = strHEDKMK & va2KOMK36 & ","             '加入員番号                         (項番36)
                                strHEDKMK = strHEDKMK & va2KOMK37 & ","             '加入形態                           (項番37)
                                strHEDKMK = strHEDKMK & va2KOMK38 & ","             '入社年月日（元号）                 (項番38)
                                strHEDKMK = strHEDKMK & va2KOMK39 & ","             '入社年月日（年月日）               (項番39)
                                strHEDKMK = strHEDKMK & va2KOMK40 & ","             '加算適用の有無                     (項番40)
                                strHEDKMK = strHEDKMK & va2KOMK41 & ","             '適用形態（取得事由コード）         (項番41)
                                strHEDKMK = strHEDKMK & va2KOMK42 & ","             '加算給与月額                       (項番42)
                                strHEDKMK = strHEDKMK & va2KOMK43 & ","             '標準給与月額                       (項番43)
                                strHEDKMK = strHEDKMK & va2KOMK44 & ","             '第２加算給与月額                   (項番44)
                                strHEDKMK = strHEDKMK & va2KOMK45 & ","             '第２加算標準給与月額               (項番45)
                                strHEDKMK = strHEDKMK & va2KOMK46 & ","             '基金固有項目１                     (項番46)
                                strHEDKMK = strHEDKMK & va2KOMK47 & ","             '基金固有項目２                     (項番47)
                                strHEDKMK = strHEDKMK & va2KOMK48 & ","             '基金固有項目３                     (項番48)
                                strHEDKMK = strHEDKMK & va2KOMK49 & ","             '基金固有項目４                     (項番49)
                                strHEDKMK = strHEDKMK & va2KOMK50 & ","             '基金固有項目５                     (項番50)
                                strHEDKMK = strHEDKMK & va2KOMK51 & ","             '基金固有項目６                     (項番51)
                                strHEDKMK = strHEDKMK & va2KOMK52 & ","             '基金固有項目７                     (項番52)
                                strHEDKMK = strHEDKMK & va2KOMK53 & ","             '基金固有項目８                     (項番53)
                                strHEDKMK = strHEDKMK & va2KOMK54 & ","             '基金固有項目９                     (項番54)
                                strHEDKMK = strHEDKMK & va2KOMK55                   '基金固有項目１０                   (項番55)

                                StWrite.WriteLine(strHEDKMK) 'CSVの書込み

                                Wcnt = Wcnt + 1
                                '[frmDBSerchG]
                                'pubDBSerchG_f.MinimizeBox = False
                                'pubDBSerchG_f.MaximizeBox = False
                                pubProgressMAX = valCOUNT9 'バーの表示
                                pubProgressCNT = CInt(Wcnt)
                                Call DBSerchG_Max(0, valCOUNT9)
                                Call DBSerchG_Value(CInt(Wcnt))

                                If pubCancelFlag = True Then 'キャンセルが押された
                                    MSGGuide(12, "データの抽出が", "", priPGMTTL)
                                    Exit While
                                End If
                            End While
                            'MessageBox.Show(valCOUNT1 & "件を読み込みました。")
                            'ループ終了

                        Case -1
                            pubUPDTFLG = sysDATAFIND.NG
                            MSGGuide(9002, "データが存在しません。", "", priPGMTTL)

                        Case -2 'タイムアウト-ロック待ち
                            'MSGGuide(17, "会社コード" & Format(valKAICD, "000") & "は", "", HKDPRGNM)
                            'TxtKAICD.SelectAll()
                            'TxtKAICD.Focus()
                        Case Else
                            pubUPDTFLG = sysDATAFIND.ER
                    End Select
                Catch ex As Exception
                    pubUPDTFLG = sysDATAFIND.ER
                    MSGGuide(9001, "(Message)" & ex.Message, vbCrLf & "(Source)" & ex.Source, "-ReadSHIKYU")
                Finally
                    sqlClosDtRdr(sysSQLDTRD)
                    If pubUPDTFLG = sysDATAFIND.ER Then Call Final()
                    Cursor.Current = Cursors.Default
                End Try
                '算定月変ファイルの読み込み終了---
                '算定月変ファイルの読み込み終了---
                '算定月変ファイルの読み込み終了---

                StWrite.Close() 'CSVﾌｧｲﾙのｸﾛｰｽﾞ

                Call DBSerchG_Close() '[frmDBSerchG]
                MSGGuide(25, valCOUNT1 & "件を出力しました。（ＣＳＶファイル）　", "", priPGMTTL)

            Else 'キャンセルが押された
                PutCSV_FDCRE3 = 9
            End If

        Catch ex As Exception
            PutCSV_FDCRE3 = 9
            MSGGuide(9001, "(Message)" & ex.Message, vbCrLf & "(Source)" & ex.Source, "-PutCSV_FDCRE3")
        End Try
    End Function

    '-------------------------------------------------------------------
    ' Form       Name : KyuMST103_f3.vb
    ' Purocedure Name : 対象件数の取得
    ' Function        : GetSSKGET_COUNT
    ' Parameter       : valKAISHA (I)   : Integer : 会社コード
    '                 : valNENGET (I)   : String  : 年月
    ' Return          : GetSSKGET_COUNT : Long    : 対象件数
    ' Date            : 2020/12/14
    '-------------------------------------------------------------------
    Public Function GetSSKGET_COUNT(ByVal valKAISHA As Integer, _
                                    ByVal valNENG_F As String, _
                                    ByVal valNENG_T As String) As Long
        Dim strSQL As String
        Dim objSqlCnnct As SqlClient.SqlConnection = Nothing 'DataAdapterで使用するConnectionｵﾌﾞｼﾞｪｸﾄを作成
        Dim objSqlDAPT As New SqlClient.SqlDataAdapter()     'DataAdapterｵﾌﾞｼﾞｪｸﾄを作成
        Dim objSqlDtSet As DataSet = New DataSet             'DataSetを作成
        Dim objSqlPara As New SqlClient.SqlParameter         'SqlParameter

        GetSSKGET_COUNT = 0

        Try
            'DataAdapterを使用するための接続
            sqlDAPCnect(objSqlCnnct)

            strSQL = ""
            strSQL = strSQL & "select count(*) as RecCnt"
            strSQL = strSQL & "  from SSKGET_DAT"
            strSQL = strSQL & " where KAISHA=@KAISHA"
            strSQL = strSQL & "   and KENNGT BETWEEN @NENG_F AND @NENG_T"

            '選択行を処理するためのCommandｵﾌﾞｼﾞｪｸﾄを作成し,SelectCommandﾌﾟﾛﾊﾟﾃｨの値に設定
            sqlDAPCommd("select", strSQL, objSqlDAPT, objSqlCnnct)

            'パラメータの設定
            sqlDAPPara("@KAISHA", SqlDbType.Int, 4, sqlPDIR.Input, valKAISHA, "select", objSqlPara, objSqlDAPT)
            sqlDAPPara("@NENG_F", SqlDbType.VarChar, 8, sqlPDIR.Input, valNENG_F, "select", objSqlPara, objSqlDAPT)
            sqlDAPPara("@NENG_T", SqlDbType.VarChar, 8, sqlPDIR.Input, valNENG_T, "select", objSqlPara, objSqlDAPT)

            'DataSourceからDataSetにﾃﾞｰﾀを読み込む
            sqlDAPFill(objSqlDAPT, objSqlDtSet, "GetSSKGET_COUNT")

            'DataTableから値を取得
            If objSqlDtSet.Tables("GetSSKGET_COUNT").Rows.Count > 0 Then
                For Each objRow In objSqlDtSet.Tables("GetSSKGET_COUNT").Rows
                    GetSSKGET_COUNT = sqlDAPFldGet(objRow("RecCnt")) '対象件数
                Next
            Else
                GetSSKGET_COUNT = 0
            End If
        Catch ex As Exception
            GetSSKGET_COUNT = -1
            MSGGuide(14, "GetSSKGET_COUNTで", "", "-GetSSKGET_COUNT")
        Finally
            'オブジェクトのクローズ
            Call sqlDAPClosDS(objSqlDtSet)
            Call sqlDAPClosDA(objSqlDAPT)
            Call sqlDAPClosCN(objSqlCnnct)
        End Try
    End Function

    '-------------------------------------------------------------------
    ' Purocedure Name : Read_KENNGT_COMB
    ' Function        : 取得年月日のコンボボックス設定
    ' Parameter       : parKAISHA (I)    : Integer : 会社コード
    ' Return          : Read_KENNGT_COMB : Short   :  0=正常
    '                                              : -1=異常
    ' Date            : 2020/12/22
    '-------------------------------------------------------------------
    Private Function Read_KENNGT_COMB(ByVal parKAISHA As Integer) As Short
        Dim strSQL As String
        'DataAdapterで使用するConnectionｵﾌﾞｼﾞｪｸﾄを作成
        Dim objSqlCnnct As SqlClient.SqlConnection = Nothing
        'DataAdapterｵﾌﾞｼﾞｪｸﾄを作成
        Dim objSqlDAPT As New SqlClient.SqlDataAdapter()
        'DataSetを作成
        Dim objSqlDtSet As DataSet = New DataSet
        'SqlParameter    
        Dim objSqlPara As New SqlClient.SqlParameter

        Me.TargetNENG_F.Clear()
        Me.TargetNENG_T.Clear()

        Read_KENNGT_COMB = 0
        Try
            'DataAdapterを使用するための接続
            sqlDAPCnect(objSqlCnnct)

            strSQL = ""
            strSQL = strSQL & "select distinct KENNGT"
            strSQL = strSQL & "  from SSKGET_DAT"
            strSQL = strSQL & " where KAISHA=@KAISHA"
            strSQL = strSQL & " order by KENNGT DESC"

            '選択行を処理するためのCommandｵﾌﾞｼﾞｪｸﾄを作成し,SelectCommandﾌﾟﾛﾊﾟﾃｨの値に設定
            sqlDAPCommd("select", strSQL, objSqlDAPT, objSqlCnnct)

            'パラメータの設定
            sqlDAPPara("@KAISHA", SqlDbType.Int, 4, sqlPDIR.Input, parKAISHA, "select", objSqlPara, objSqlDAPT)

            'DataSourceからDataSetにﾃﾞｰﾀを読み込む
            sqlDAPFill(objSqlDAPT, objSqlDtSet, "Read_KENNGT_COMB")

            'DataTableから値を取得
            If objSqlDtSet.Tables("Read_KENNGT_COMB").Rows.Count > 0 Then
                For Each objRow In objSqlDtSet.Tables("Read_KENNGT_COMB").Rows
                    Dim valTARGET As String = sqlDAPFldGet(objRow("KENNGT"))
                    Dim lngTARGET As Long = Long.Parse(valTARGET)
                    If lngTARGET > 0 Then
                        Me.TargetNENG_F.Add(DateTime.ParseExact(valTARGET, "yyyyMMdd", Nothing))
                        Me.TargetNENG_T.Add(DateTime.ParseExact(valTARGET, "yyyyMMdd", Nothing))
                    End If
                Next
            Else
                Read_KENNGT_COMB = -1
            End If
        Catch ex As Exception
            Read_KENNGT_COMB = 9
            MSGGuide(14, "Read_KENNGT_COMBで", "", "-Read_KENNGT_COMB")
        Finally
            'オブジェクトのクローズ
            Call sqlDAPClosDS(objSqlDtSet)
            Call sqlDAPClosDA(objSqlDAPT)
            Call sqlDAPClosCN(objSqlCnnct)
        End Try
    End Function


    '-------------------------------------------------------------------
    ' Purocedure Name : ConvertDate_Fnc
    ' Function        : 和暦を西暦に変換　平成21年12月12日→20091212
    ' Parameter       : valWAREKI (I)     : String : 和暦　平成21年12月12日
    ' Return          : ConvertDate_Fnc   : String : 西暦　20091212
    ' Date            : 2009/11/18
    '-------------------------------------------------------------------
    Public Function ConvertDate_Fnc(ByVal valWAREKI As String) As String
        ConvertDate_Fnc = ""

        Dim strWAREKI As String = valWAREKI
        strWAREKI = Replace(strWAREKI, " ", "")

        '元号の取得
        Dim valGENGOU As String = Strings.Left(strWAREKI, 2)
        Select Case valGENGOU
            Case "明治" : valGENGOU = "1"
            Case "大正" : valGENGOU = "2"
            Case "昭和" : valGENGOU = "3"
            Case "平成" : valGENGOU = "4"
            Case "令和" : valGENGOU = "5"
        End Select

        'YYMMDDの取得
        Dim valYYMMDD As String = Strings.Right(strWAREKI, 9)
        valYYMMDD = Replace(valYYMMDD, "年", "")
        valYYMMDD = Replace(valYYMMDD, "月", "")
        valYYMMDD = Replace(valYYMMDD, "日", "")

        '和暦を西暦
        valYYMMDD = valGENGOU & Strings.Format(CInt(valYYMMDD), "000000")
        valYYMMDD = ConvertDate(1, valYYMMDD)
        ConvertDate_Fnc = valYYMMDD
    End Function
End Class
