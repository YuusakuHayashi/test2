Attribute VB_Name = "Rpa01"

Option Base 0

' 課題
'   1. save_list() のAdd2 が動かない？
'      verの関係上？

Private Enum layout_type
    NORMAL = 0
    NORMAL2 = 1
    GNET = 2
    YUCHO = 3
    ASIKAGA = 4
End Enum

Private Enum sheet_change
    GINKOU = 0
    SITEN = 1
End Enum

Private Const FSOBJ            As String = "Scripting.FileSystemObject"
Private Const WORK_XLSX        As String = "work.xlsx"
Private Const RPA_XLM          As String = "rpa.xlsm"
Private Const TEISHI_LIST      As String = "停止リスト.xls"
Private Const IRAISHO_GNET1    As String = "停止依頼書(G-NET1).xls"
Private Const IRAISHO_GNET2    As String = "停止依頼書(G-NET2).xls"
Private Const IRAISHO_GNET3    As String = "停止依頼書(G-NET3).xls"
Private Const IRAISHO_SIS      As String = "停止依頼書(SIS).xlsx"
Private Const IRAISHO_SIS_ZIEI As String = "停止依頼書(SIS自営信金).xlsx"
Private Const IRAISHO_MIZUHO   As String = "停止依頼書(みずほ銀行).xls"
Private Const IRAISHO_YUCHO    As String = "停止依頼書(ゆうちょ銀行).xls"
Private Const IRAISHO_MITSUI   As String = "停止依頼書(三井住友銀行).xls"
Private Const IRAISHO_TAIKO    As String = "停止依頼書(大光銀行).xls"
Private Const IRAISHO_JOYO     As String = "停止依頼書(常陽銀行).xls"
Private Const IRAISHO_AKITA    As String = "停止依頼書(秋田銀行).xls"
Private Const IRAISHO_ASIKAGA  As String = "停止依頼書(足利銀行).xls"
Private Const IRAISHO_NORMAL   As String = "停止依頼書.xls"

Private Type runtime
    work_book As String
End Type

Private Type list_data
    book_name           As String
    book_obj            As Workbook
    sheet_obj           As Worksheet
    primary_data_column As Integer
    last_data_row       As Integer
    ginkou_name         As String
    siten_name          As String
    kouza_shubet        As String
    kouza_meigi         As String
    kingaku             As String
    kokyaku_code        As String
    touroku_name        As String
    ginkou_name_cell    As Range
    siten_name_cell     As Range
    kouza_shubet_cell   As Range
    kouza_meigi_cell    As Range
    kingaku_cell        As Range
    kokyaku_code_cell   As Range
    touroku_name_cell   As Range
    output_csv_name     As String
End Type

Private Type meisai_data
'   HISTORY
'      2020/08/13: 自営信金コード追加
    ginkou_name   As String
    siten_name    As String
    kouza_shubet  As String
    kouza_num     As String
    kouza_meigi   As String
    kingaku       As String
    bikou         As String
    ginkou_code   As String
    siten_code    As String
    furikae_date  As String
    tucho_num     As String
    ginkou_code1  As String
    ginkou_code2  As String
    ginkou_code3  As String
    ginkou_code4  As String
    ginkou_saffix As String
    kokyaku_code  As String
    funo_code     As String
    ziei_code     As String
End Type

Private Type summary_data
'   HISTORY
'      2020/02/01: 新規作成
    ginkou_name   As String
    data_count    As String
End Type

Private Type rpa_project
    json_name   As String
    project_dir As String
    runtime_dir As String
End Type

Private Type iraisho_data
'   HISTORY
'      2020/01/28: start_code     追加
'      2020/01/28: start_flag     追加
'      2020/08/13: ziei_code_cell 追加
'
    start_code          As String
    start_flag          As Boolean
    writer_name         As String
    config_file_name    As String
    sheet_layout        As Integer
    genpon_sheet_name   As String
    genpon_sheet_obj    As Worksheet
    write_sheet_name    As String
    write_sheet_obj     As Worksheet
    book_name           As String
    book_obj            As Workbook
    sheet_m             As String
    meisai_d            As meisai_data
    sheet_change_type   As Integer
    write_row           As Integer
    max_write_row       As Integer
    '
    writer_name_cell    As Range
    ginkou_name_cell    As Range
    siten_name_cell     As Range
    kouza_shubet_cell   As Range
    kouza_num_cell      As Range
    kouza_meigi_cell    As Range
    kingaku_cell        As Range
    bikou_cell          As Range
    ginkou_code_cell    As Range
    siten_code_cell     As Range
    furikae_date_cell   As Range
    tucho_num_cell      As Range
    ginkou_code1_cell   As Range
    ginkou_code2_cell   As Range
    ginkou_code3_cell   As Range
    ginkou_code4_cell   As Range
    ginkou_saffix_cell  As Range
    kokyaku_code_cell   As Range
    funo_code_cell      As Range
    ziei_code_cell      As Range
    primary_data_column As Integer
    '
    project             As rpa_project
End Type

Private Function convert_to_omoji(ByVal moji As String)
    moji = Replace(moji, "ァ", "ア")
    moji = Replace(moji, "ィ", "イ")
    moji = Replace(moji, "ゥ", "ウ")
    moji = Replace(moji, "ェ", "エ")
    moji = Replace(moji, "ォ", "オ")
    moji = Replace(moji, "ャ", "ヤ")
    moji = Replace(moji, "ュ", "ユ")
    moji = Replace(moji, "ョ", "ヨ")
    moji = Replace(moji, "ッ", "ツ")
    moji = Replace(moji, "ｧ", "ｱ")
    moji = Replace(moji, "ｨ", "ｲ")
    moji = Replace(moji, "ｩ", "ｳ")
    moji = Replace(moji, "ｪ", "ｴ")
    moji = Replace(moji, "ｫ", "ｵ")
    moji = Replace(moji, "ｬ", "ﾔ")
    moji = Replace(moji, "ｭ", "ﾕ")
    moji = Replace(moji, "ｮ", "ﾖ")
    moji = Replace(moji, "ｯ", "ﾂ")
    convert_to_omoji = moji
End Function

Private Property Get SYS_DIR() As String
'   HISTORY
'      2020/03/26: 新規作成
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    SYS_DIR = wsh.ExpandEnvironmentStrings("%USERPROFILE%") + "\" + "rpa_project"
    Set wsh = Nothing
End Property

Private Property Get SYSTEM_DIRECTORY() As String
'   HISTORY
'      2020/03/26: 新規作成
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    SYSTEM_DIRECTORY = wsh.ExpandEnvironmentStrings("%USERPROFILE%") + "\" + "rpa_project"
    Set wsh = Nothing
End Property

Private Property Get SYS_SYS() As String
'   HISTORY
'      2020/03/26: 新規作成
    SYS_SYS = SYS_DIR + "\" + "sys"
End Property

'Private Property Get PRJ_JSN As String
''   HISTORY
''      2020/03/26: 新規作成
''      2020/04/28: ディレクトリ構成の変更による変更
'    PRJ_JSN = SYS_SYS + "\" + "script" + "\" + "rpa_project.json"
'End Property

Private Property Get SystemJsonFileName() As String
'   HISTORY
'      2020/03/26: 新規作成
'      2020/04/28: ディレクトリ構成の変更による変更
    SystemJsonFileName = SYSTEM_DIRECTORY + "\rpa_project.json"
End Property

'Private Function GetProjectJson() As Object
'    Dim fso As Object
'    Dim txs As Object
'    Dim txt As String
'    Dim jsn As Object
'    Set fso = CreateObject(FSOBJ)
'    Set txs = fso.OpenTextFile(SystemJsonFileName, 1, -2)
'    txt = RTrim(txs.ReadAll)
'    Set jsn = JsonConverter.ParseJson(txt)
'    Set GetProjectJson = jsn
'    txs.Close
'    Set txs = Nothing
'    Set fso = Nothing
'    Set jsn = Nothing
'End Function

' Jsonファイルの読み込み(UTF-8)
'--------------------------------------------------------------------'
Private Function GetProjectJson() As Object
    Dim ado As Object
    Dim txt As String
    Dim json As Object

    Set ado = CreateObject("ADODB.Stream")
    ado.Charset = "UTF-8"
    Call ado.Open
    Call ado.LoadFromFile(SystemJsonFileName)
    txt = RTrim(ado.ReadText)
    Set json = JsonConverter.ParseJson(txt)
    Set GetProjectJson = json
    Call ado.Close

    Set json = Nothing
    Set ado = Nothing
End Function
'--------------------------------------------------------------------'

Private Sub set_print_setting()
    Application.PrintCommunication = False
    Dim sh As Object
    For Each sh In Worksheets
        If InStr(sh.Name, "停止") > 0 Then
            With sh.PageSetup
                .LeftMargin = Application.InchesToPoints(0.236220472440945)
                .RightMargin = Application.InchesToPoints(0.236220472440945)
                .TopMargin = Application.InchesToPoints(0.748031496062992)
                .BottomMargin = Application.InchesToPoints(0.748031496062992)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .CenterHorizontally = False
                .CenterVertically = False
                .Orientation = xlLandscape
                .FitToPagesWide = 1
                .FitToPagesTall = 1
            End With
        End If
    Next
    Application.PrintCommunication = True
End Sub

Private Sub save_master_as_xlsx()
    Call save_csv_as_xlsx(1)
End Sub

Private Sub save_csv_as_xlsx(ByRef mode As Integer)
'   NOTE
'      project_dir で実行
'   HISTORY
'      2020/01/28: 途中開始する対応
'      2020/03/14: work.xlsm対応
'--- 途中開始対応 ----------------------------------------------*
'   Dim start_code As String: start_code = get_file_text(ThisWorkbook.Path + "\" + "start_code")
'
'   If Not start_code = "XXXXXXXXXXXXX" Then
'       Exit Sub
'   End If
'---------------------------------------------------------------*
    Dim wb As Workbook
    Dim in_file As String
    Dim jsn As Object: Set jsn = GetProjectJson()

    Select Case mode
        Case 1
            in_file = jsn("USR_PRJ_WRK") + "\" + "master.csv"
        Case 2
            in_file = jsn("USR_PRJ_WRK") + "\" + "teishi_result.csv"
        Case 3
            in_file = jsn("USR_PRJ_WRK") + "\" + "tmp.csv"
    End Select

    Set wb = Workbooks.Open(in_file, 0)
    With wb
        .Application.DisplayAlerts = False
        Dim n As String: n = .Name
        n = Replace(n, ".csv", "")
        n = Replace(n, ".CSV", "")
        .SaveAs Filename:=jsn("USR_PRJ_WRK") + "\" + "modified_" + n, FileFormat:=Excel.xlWorkbookDefault
    End With
    wb.Close
    Set wb = Nothing
    Set jsn = Nothing
End Sub

Private Sub check_list(ByRef wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.ActiveSheet
    Dim rngs As Range
    Dim rng As Range
    Set rngs = ws.Cells(1, 1).CurrentRegion
    For Each rng In rngs.Rows
        Select Case rng.Columns(20).Text
            Case vbNullString
            Case 15
                rng.Interior.ColorIndex = 6
            Case Else
        End Select
    Next
    Set rngs = Nothing
    Set ws = Nothing
End Sub

Private Sub sort_teishi_list()
'   HISTORY
'      2020/03/16: 新規作成
    Dim wrk_bok As Workbook
    Dim wsh     As Worksheet
    Dim lst_row As Integer
    Dim jsn As Object: Set jsn = GetProjectJson()
    Set wrk_bok = Workbooks.Open(jsn("USR_PRJ_WRK") + "\" + "input.xls", 0)

    wrk_bok.Application.EnableEvents = False
    wrk_bok.Application.DisplayAlerts = False

    lst_row = wrk_bok.ActiveSheet.Cells(Cells.Rows.Count, 1).End(xlUp).row

    Set wsh = wrk_bok.ActiveSheet

    With wsh.AutoFilter.Sort
        With .SortFields
            .Clear
            .Add Key:=Range("C13"), _
                 Order:=xlAscending, _
                 DataOption:=xlSortNormal
        End With
        .Apply
    End With

    wrk_bok.Save
    wrk_bok.Close

    Set wsh = Nothing
    Set wrk_bok = Nothing
    Set jsn = Nothing
End Sub

Private Sub save_list()
'   HISTORY
'      2020/01/28: 途中開始する対応

    Dim last_row As Integer
'--- 途中開始対応 ----------------------------------------------*
    Dim distination_book As Workbook
'---------------------------------------------------------------*
    Dim sheet_name As String: sheet_name = "停止"
    Dim jsn As Object: Set jsn = GetProjectJson()

    Call save_csv_as_xlsx(2)

    Dim source_book As Workbook
    Set source_book = Workbooks.Open( _
        jsn("USR_PRJ_WRK") + "\" + "modified_teishi_result.xlsx", 0 _
    )
    source_book.Application.DisplayAlerts = False
    source_book.Application.EnableEvents = False

'--- 途中開始対応 ----------------------------------------------*
    Set distination_book = Workbooks.Open( _
        jsn("USR_PRJ_WRK") + "\" + "modified_master.xlsx", 0 _
    )

    distination_book.Application.DisplayAlerts = False
    distination_book.Application.EnableEvents = False

    Dim wsh_cnt As Integer: wsh_cnt = distination_book.Worksheets.Count
    'For i = 1 To distination_book.Worksheets.Count
    '    If distination_book.Worksheets(i).Name = sheet_name Then
    '        sheet_name = sheet_name + "I"
    '    End If
    'Next
    sheet_name = sheet_name + String(wsh_cnt - 1, "I")
    source_book.ActiveSheet.Name = sheet_name
'---------------------------------------------------------------*

    Call check_list(source_book)

    With source_book.ActiveSheet
        .Columns("A:B").Hidden = True
        .Columns("F:G").Hidden = True
        .Columns("Q:U").Hidden = True
        .Columns("E").AutoFit
        .Columns("M:N").AutoFit
    End With

    last_row = source_book.ActiveSheet.Cells(Cells.Rows.Count, 8).End(xlUp).row

    If last_row > 1 Then
        'With source_book.WorkSheets(sheet_name).AutoFilter.Sort
        '    With .SortFields
        '        .Clear
        '        .Add Key:=Range("H1"), _
        '             Order:=xlAscending, _
        '             DataOption:=xlSortNormal
        '    End With
        '    .Apply
        'End With
        'source_book.WorkSheets(sheet_name).Cells(2, 3).Select
        'source_book.Application.ActiveWindow.FreezePanes = True
        With source_book.Worksheets(sheet_name)
            .Range("H1").AutoFilter
            .AutoFilter.Sort.SortFields.Clear
            .AutoFilter.Sort.SortFields.Add _
                Key:=Range(Cells(1, 8), Cells(last_row, 8)), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .AutoFilter.Sort.Apply
            .Activate
            .Cells(2, 3).Select
            source_book.Application.ActiveWindow.FreezePanes = True
        End With
    End If
    source_book.Save

    source_book.Worksheets(sheet_name).Copy after:= _
        distination_book.Worksheets(distination_book.Worksheets.Count)

    source_book.Close
    distination_book.Save
    distination_book.Close

    Set distination_book = Nothing
    Set source_book = Nothing
    Set jsn = Nothing
End Sub

Private Function get_file_text(txt_file) As String: get_file_text = vbNullString
    Dim fso As Object
    Set fso = CreateObject(FSOBJ)
    Dim s As String
    Dim txt As String
    With fso
        Dim ts As Object
        Set ts = fso.OpenTextFile(txt_file, 1, -2)
        With ts
            get_file_text = Trim(ts.ReadLine())
            .Close
        End With
        Set ts = Nothing
    End With
    Set fso = Nothing
End Function

Private Sub check_main()
    Call write_check_list_csv
    Dim fso As Object
    Set fso = CreateObject(FSOBJ)
    With fso
        Dim ts As Object
        Set ts = fso.OpenTextFile(get_file_text(ThisWorkbook.Path + "\runtime_dir"), 1, -2)
        With ts
            .Close
        End With
        Set ts = Nothing
    End With
    Set fso = Nothing
End Sub

'Private Function write_check_list() As String()
Private Sub write_check_list_csv()
    Dim wss As Sheets: Set wss = ThisWorkbook.Worksheets
    Dim ws As Worksheet
    Dim rngs As Range
    Dim rng As Range
    Dim r As Range
    Dim fso As Object
    Set fso = CreateObject(FSOBJ)
    Dim jsn As Object: Set jsn = GetProjectJson()
    Dim s As String
    Dim txt As String
    'Dim vs(0 To 100, 0 To 6) As String
    With fso
    Dim ts As Object
    Set ts = fso.OpenTextFile(jsn("USR_PRJ_WRK") + "\check_list.csv", 2, -2)
    'Dim x As Integer: x = 0
    'Dim y As Integer: y = 0
    For Each ws In wss
        If InStr(ws.Name, "NO") Then
            Set rngs = ws.Range("B29:AL37")
            For Each rng In rngs.Rows
                For Each r In rng.Columns
                    s = r.Text
                    If Not s = vbNullString Then
                        If Information.IsNumeric(s) Then
                            s = Replace(s, ",", "")
                        End If
                        'vs(x, y) = s
                        'y = y + 1
                        If txt = vbNullString Then
                            txt = StrConv(s, vbNarrow)
                        Else
                            txt = txt + "," + StrConv(s, vbNarrow)
                        End If
                    End If
                Next
                'y = 0
                'x = x + 1
                ts.WriteLine (txt)
                txt = vbNullString
            Next
        End If
    Next
    'get_check_datas = vs
    ts.Close
    Set ts = Nothing
    End With
    Set r = Nothing
    Set rng = Nothing
    Set rngs = Nothing
    Set fso = Nothing
    Set jsn = Nothing
End Sub

Public Sub xl_to_csv()
    Dim fso As Object
    Dim f As Object
    Set fso = CreateObject(FSOBJ)
    Set f = fso.OpenTextFile("tmp.csv", 2, -2)
    Dim rng As Range
    Set rng = Cells(1, 1).CurrentRegion
    Dim maxrow As Integer: maxrow = rng.Rows.Count
    Dim maxcol As Integer: maxcol = rng.Columns.Count
    Dim txt As String
    Dim i As Integer
    Dim j As Integer
    For i = 1 To maxrow
        txt = rng(i, 1).Text
        For j = 2 To maxcol
            txt = txt + "," + rng(i, j).Text
        Next
        If rng(i, 3).Text = "29005" Then
            f.WriteLine (txt)
        End If
        txt = ""
    Next
    f.Close
    Set rng = Nothing
    Set f = Nothing
    Set fso = Nothing
End Sub

Sub clear_current_reigion()
    Range("A1").CurrentRegion.ClearContents
End Sub

Private Sub extract_csv(ParamArray pa() As Variant)
    Dim additional_flag As Boolean: additional_flag = pa(0)
    Dim data_hit_flag As Boolean: data_hit_flag = pa(1)
    Dim header_data As String: header_data = pa(2)
    Dim source_data As String: source_data = pa(3)
    If data_hit_flag Then
        Dim wb As Workbook
        With Application
            .DisplayAlerts = False
            .EnableEvents = False
            Set wb = Workbooks.Open(ThisWorkbook.Path + "\NEW_KOKMAT.xlsx", 0)
            With wb
                .Application.DisplayAlerts = False
                .Application.EnableEvents = False
                With .ActiveSheet
                    Dim datas() As String
                    Dim i As Integer
                    Dim r As Integer: r = 1
                    Dim c As Integer: c = 1
                    If Not additional_flag Then
                        datas = Split(header_data, ",")
                        For i = LBound(datas) To UBound(datas)
                            .Cells(r, c) = datas(i)
                            c = c + 1
                        Next
                    End If
                    datas = Split(source_data, ",")
                    r = .Cells(Rows.Count, 1).End(xlUp).Offset(1).row
                    c = 1
                    For i = LBound(datas) To UBound(datas)
                        .Cells(r, c) = datas(i)
                        c = c + 1
                    Next
                End With
                .Save
                .Close
'               .Application.DisplayAlerts = True
'               .Application.EnableEvents = True
            End With
            Set wb = Nothing
'           .DisplayAlerts = True
'           .EnableEvents = True
        End With
    End If
End Sub

Private Function get_list_of_file(ByRef file_name As String) As String()
    Dim fso As Object: Set fso = Nothing
    Dim file_list As String: file_list = file_name
    Set fso = CreateObject(FSOBJ)
    With fso
        Dim ts As Object: Set ts = Nothing
        Set ts = .OpenTextFile(file_name, 1, True, -2)
        Dim i As Integer: i = 0
        Dim tmp() As String
        ReDim tmp(100)
        Do Until ts.AtEndOfStream
            tmp(i) = ts.ReadLine
            i = i + 1
        Loop
        ReDim Preserve tmp(i - 1)
        get_list_of_file = tmp
        Set ts = Nothing
    End With
    Set fso = Nothing
End Function

Private Property Get project_directory() As String
'   HISTORY
'      2020/02/01: 新規作成
    project_directory = get_file_text(ThisWorkbook.Path + "\project_dir")
End Property

Private Property Get runtime_directory() As String
'   HISTORY
'      2020/02/01: 新規作成
    runtime_directory = get_file_text(ThisWorkbook.Path + "\runtime_dir")
End Property

Private Sub delete_sheet()
'   NOTE
'      runtime_directory で実行
'      未使用と思われる
'   HISTORY
'      2019/08/22: 削除ロジックをクリーンにした
'      2019/10/25: 対象エクセルファイルの外部化
'      2020/01/28: 途中開始する対応
'      2020/08/13: JSON対応
'-- 途中開始対応--------------------------------------------------*
    Dim start_code As String: start_code = get_file_text(ThisWorkbook.Path + "\" + "start_code")

    If Not start_code = "XXXXXXXXXXXXX" Then
        Exit Sub
    End If
'-----------------------------------------------------------------*
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim tgt As Integer
    Dim wb As Workbook

    Dim jsn As Object: Set jsn = GetProjectJson()
    Dim wbs As Object: Set wbs = jsn("ROO_PRJ_RPA")("iraisyo")
    'Dim file_name As String: file_name = ThisWorkbook.Path + "\iraisyo"
    'Dim wbs() As String: wbs = get_list_of_file(file_name)
    With Application
        .DisplayAlerts = False
        .EnableEvents = False
        'For k = LBound(wbs) To UBound(wbs)
        For Each wbn In wbs
            Set wb = Workbooks.Open( _
                rpa.project_directory + "\" + wbn, 0)
            With wb
                Application.DisplayAlerts = False
                Application.EnableEvents = False
                Application.ScreenUpdating = False
                'For i = .Worksheets.Count To i Step -1
                Dim wss() As String
                ReDim wss(.Worksheets.Count - 1)
                For i = LBound(wss) To UBound(wss)
                    wss(i) = .Worksheets(i + 1).Name
                Next
                
                For i = LBound(wss) To UBound(wss)
                    If InStr(1, wss(i), "原本", vbTextCompare) = 0 Then
                        Worksheets(wss(i)).Delete
                    End If
                Next
                .Save
                .Close
                Set wb = Nothing
'               .Application.DisplayAlerts = True
'               .Application.EnableEvents = True
            End With
'            .DisplayAlerts = True
'            .EnableEvents = True
        Next
    End With

    Set wbs = Nothing
    Set jsn = Nothing
End Sub

Private Sub print_summary()
'   NOTE
'      runtime_directory で実行
'   HISTORY
'      2020/02/04: 新規作成
    Dim wb As Workbook
    Dim jsn As Object: Set jsn = GetProjectJson()
    Set wb = Workbooks.Open( _
        jsn("USR_PRJ_WRK") + "\" + "summary.csv", 0)
    With wb
        .Application.DisplayAlerts = False
        .Application.EnableEvents = False
        .Application.ScreenUpdating = False
        .ActiveSheet.Columns("A:B").AutoFit
        .ActiveSheet.PrintOut
        .Close
    End With
    Set wb = Nothing
    Set jsn = Nothing
End Sub

Private Sub print_teishi_list()
'   NOTE
'      runtime_directory で実行
'   HISTORY
'      2020/02/04: 新規作成
'      2020/03/14: work.xlsm対応
    Dim wb As Workbook
    Dim jsn As Object: Set jsn = GetProjectJson()
    Set wb = Workbooks.Open(jsn("USR_PRJ_WRK") + "\" + "input.xls", 0)
    lst_row = wb.ActiveSheet.Cells(Cells.Rows.Count, 1).End(xlUp).row
    Dim i As Integer: i = 0
    For i = lst_row + 5 To 160
        wb.ActiveSheet.Rows(i).Hidden = True
    Next
    With wb
        .Application.DisplayAlerts = False
        .Application.EnableEvents = False
        .Application.ScreenUpdating = False
        .ActiveSheet.PrintOut
        .Close
    End With
    Rows.Hidden = False
    Set wb = Nothing
    Set jsn = Nothing
End Sub

Private Sub print_soufu_meisai()
'   NOTE
'      runtime_directory で実行
'   HISTORY
'      2020/02/04: 新規作成
'      2020/03/14: work.xlsm対応
    Dim wb As Workbook
    Dim jsn As Object: Set jsn = GetProjectJson()
    Set wb = Workbooks.Open( _
        jsn("USR_PRJ_WRK") + "\" + "modified_master.xlsx", 0 _
    )
    With wb
        .Application.DisplayAlerts = False
        .Application.EnableEvents = False
        .Application.ScreenUpdating = False
        For i = 1 To .Worksheets.Count
            If InStr(1, .Worksheets(i).Name, "停止", vbTextCompare) = 1 Then
                .Worksheets(i).PageSetup.Orientation = xlLandscape
                .Worksheets(i).PrintOut
            End If
        Next
        .Close
    End With
    Set wb = Nothing
    Set jsn = Nothing
End Sub

Private Sub print_sheet()
'   NOTE
'      runtime_directory で実行
'   HISTORY
'      2020/02/04: 新規作成
'      2020/03/14: work.xlsm対応
'      2020/03/27: jsn対応
'      2020/08/13: jsn対応のバグ修正
    Dim wb As Workbook
    Dim jsn As Object: Set jsn = GetProjectJson()
    'Dim file_name As String: file_name = jsn("USR_PRJ_WRK") + "\iraisyo"
    'Dim wbs() As String: wbs = get_list_of_file(file_name)
    Dim wbs As Object: Set wbs = jsn("ROO_PRJ_RPA")("iraisyo")

    With Application
        .DisplayAlerts = False
        .EnableEvents = False
        'For k = LBound(wbs) To UBound(wbs)
        For Each wbn In wbs
            Set wb = Workbooks.Open(jsn("USR_PRJ_WRK") + "\" + wbn, 0)
            With wb
                .Application.DisplayAlerts = False
                .Application.EnableEvents = False
                .Application.ScreenUpdating = False
                For i = 1 To wb.Worksheets.Count
                    If InStr(1, wb.Worksheets(i).Name, "原本", vbTextCompare) = 0 Then
                        Worksheets(i).PrintOut ActivePrinter:=jsn("SYS_PRJ_RPA")("printer")
                    End If
                Next
                .Close
                Set wb = Nothing
            End With
        Next
    End With
    Set wbs = Nothing
    Set jsn = Nothing
End Sub

Sub set_writer_test()
    Call set_writer("林")
End Sub

Private Sub set_writer(ParamArray pa() As Variant)
'   HISTORY
'       - 2019/10/25 : 対象エクセルファイルの外部化
'       - 2019/12/04 : 廃止
    Exit Sub
    Dim writer_name As String: writer_name = pa(0)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim tgt As Integer
    Dim wb As Workbook
    Dim file_name As String: file_name = get_file_text(ThisWorkbook.Path + "\runtime_dir") + "\iraisyo"
    Dim wbs() As String: wbs = get_list_of_file(file_name)
    With Application
        .DisplayAlerts = False
        .EnableEvents = False
        For k = LBound(wbs) To UBound(wbs)
            Set wb = Workbooks.Open(ThisWorkbook.Path + "\" + wbs(k), 0)
            With wb
                Application.DisplayAlerts = False
                Application.EnableEvents = False
                Application.ScreenUpdating = False
                'For i = .Worksheets.Count To i Step -1
                Dim wss() As String
                ReDim wss(.Worksheets.Count - 1)
                For i = LBound(wss) To UBound(wss)
                    wss(i) = .Worksheets(i + 1).Name
                Next
               
                Dim rng As Range
                Set rng = Nothing
                For i = LBound(wss) To UBound(wss)
                    If InStr(1, wss(i), "原本", vbTextCompare) <> 0 Then
                        Select Case wbs(k)
                            Case wbs(0)
                                Set rng = Worksheets(wss(i)).Range("Q8")
                            Case wbs(4)
                                If wss(i) = "原本(自営)" Then
                                    Set rng = Worksheets(wss(i)).Range("Q9")
                                Else
                                    Set rng = Worksheets(wss(i)).Range("Q8")
                                End If
                            Case wbs(5)
                                Set rng = Worksheets(wss(i)).Range("Q8")
                            Case wbs(6)
                                Set rng = Worksheets(wss(i)).Range("I12")
                            Case wbs(7)
                                Set rng = Worksheets(wss(i)).Range("Q9")
                            Case wbs(8)
                                Set rng = Worksheets(wss(i)).Range("Q8")
                            Case wbs(9)
                                Set rng = Worksheets(wss(i)).Range("Q9")
                            Case wbs(10)
                                Set rng = Worksheets(wss(i)).Range("Q8")
                        End Select
                        If Not rng Is Nothing Then
                            rng = writer_name
                        End If
                        Set rng = Nothing
                    End If
                Next
                .Save
                .Close
                Set wb = Nothing
'               .Application.DisplayAlerts = True
'               .Application.EnableEvents = True
            End With
'            .DisplayAlerts = True
'            .EnableEvents = True
        Next
    End With
End Sub

Function gbn(ByVal n As String)
    Dim fso As Object
    Set fso = CreateObject(FSOBJ)
    gbn = fso.GetBaseName(n)
    Set fso = Nothing
End Function

Function check_end_flag() As Boolean: check_end_flag = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    Dim s As String
    Dim r As Range
    Dim rng As Range
    s = ".*最終.*"
    With re
        .Pattern = s
        .IgnoreCase = False
        .Global = False
        With ws
            Set rng = Range( _
                .Range("A1"), _
                .Range("A1") _
                    .SpecialCells(xlCellTypeVisible) _
                        .SpecialCells(xlCellTypeLastCell) _
            )
        End With
        For Each r In rng
            If .test(r.Text) Then
                check_end_flag = True
            End If
        Next r
    End With
    Set rng = Nothing
    Set re = Nothing
    Set ws = Nothing
End Function

'Sub copy_workbooks(ByVal s As String, ByVal b As String)
Sub copy_workbooks()
'   HISTORY
'      2019/10/25: 対象エクセルファイルの外部化
'      2020/03/27: 使用してないと思われる
    'Dim shl As Object
    'Set shl = CreateObject("WScript.Shell")
    Dim bp As String
    Dim d As String
    Dim c As String
    Dim fso As Object
    Set fso = CreateObject(FSOBJ)

    bp = ThisWorkbook.Path & "\backup"
    If fso.FolderExists(bp) Then
    Else
        MkDir bp
    End If

    Dim i As Integer
    Dim wb As Workbook
    Dim file_name As String: file_name = get_file_text(ThisWorkbook.Path + "\runtime_dir") + "\iraisyo"
    Dim wbs() As String: wbs = get_list_of_file(file_name)
    With Application
        .DisplayAlerts = False
        .EnableEvents = False
        For i = LBound(wbs) To UBound(wbs)
            Set wb = Workbooks.Open(ThisWorkbook.Path + "\" + wbs(i), 0)
            With wb
                .Application.DisplayAlerts = False
                .Application.EnableEvents = False
                .Application.ScreenUpdating = False
                Select Case wbs(i)
                    Case IRAISHO_SIS, IRAISHO_SIS_ZIEI
                        .SaveAs (bp & "\" & gbn(.Name) + "_" + Format(Date, "yyyymmdd") & ".xlsx")
                    Case Else
                        .SaveAs (bp & "\" & gbn(.Name) + "_" + Format(Date, "yyyymmdd") & ".xls")
                End Select
                .Close
                Set wb = Nothing
                '.Application.DisplayAlerts = True
                '.Application.EnableEvents = True
            End With
        Next
'       .DisplayAlerts = True
'       .EnableEvents = True
    End With

    'b = ThisWorkbook.Path & "\backup.bat"
    'b = b & "\backup.bat"

    'd = bp & "\" & gbn(s) & "_" & Format(Date, "yyyymmdd") & ".xls"

    'c = b + " " + s + " " + d
    'Dim outcome As Long
    'outcome = shl.Run(c, WaitOnReturn:=True)
    'Set shl = Nothing
    Set fso = Nothing
End Sub

'Sub copy_me()
'    Dim wb As Workbook
'    Set wb = ThisWorkbook
'
'    Dim fso As Object
'    Set fso = CreateObject(FSOBJ)
'    Dim bp As String
'    bp = wb.Path & "\backup"
'    If fso.FolderExists(bp) Then
'    Else
'        MkDir bp
'    End If
'
'    wb.SaveCopyAs _
'        (wb.Path _
'        & "\" & "backup" _
'        & "\" & gbn(wb.Name) _
'        & "_" & Format(Date, "yyyymmdd") _
'        & ".xls")
'End Sub

'Sub delete_posfile()
'    If check_last Then
'        Dim tmp As String: tmp = ThisWorkbook.Path + "\spos.tmp"
'        'Dim tmp As String: tmp = "spos.tmp"
'        Dim fso As Object
'        Dim m As String
'        Dim tmpf As Object
'        Set fso = CreateObject(FSOBJ)
'        If fso.FileExists(tmp) Then
'            fso.DeleteFile tmp, True
'        End If
'        Set fso = Nothing
'        'Call copy_workbooks
'    End If
'End Sub

Function get_errmsg(ByRef rng As Range) As String: get_errmsg = ""
    '現在、顧客コード、請求額いずれかがない場合、True
    'リテラルで顧客コードと、請求額セル位置を判別。修正余地あり
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim ed As struct_indata: ed.errmsg = ""
    Dim r As Range
    For Each r In rng
        With r
            If .Text = "" Then
                ed.errmsg = ed.errmsg + "セル位置(" + Str(.row) + "," + Str(.Column) + "): 空白" + vbCrlf
            Else
                If .Column = 4 + 1 Or .Column = 5 + 1 Then
                    If Not IsNumeric(.Text) Then
                        ed.errmsg = ed.errmsg + "セル位置(" + Str(.row) + "," + Str(.Column) + "): 非数字" + vbCrlf
                    End If
                Else
                    If IsNumeric(.Text) Then
                        ed.errmsg = ed.errmsg + "セル位置(" + Str(.row) + "," + Str(.Column) + "): 数字" + vbCrlf
                    End If
                End If
            End If
        End With
    Next
    get_errmsg = ed.errmsg
End Function

'Sub set_errmsg(ByVal row As Integer, ByVal lastcol As Integer)
'    Dim rng As Range
'    Dim msg As String
'    Set rng = Range(Cells(row, 1), Cells(row, lastcol))
'    msg = get_errmsg(rng)
'
'    With rng
'        rng(.Rows.Count, .Columns.Count).AddComment msg
'    End With
'    Set rng = Nothing
'End Sub
'
'Sub delete_errmsg(ByVal row As Integer, ByVal lastcol As Integer)
'    Dim rng As Range
'    Set rng = Range(Cells(row, 1), Cells(row, lastcol))
'    rng(rng.Rows.Count, rng.Columns.Count).ClearComments
'    Set rng = Nothing
'End Sub

'Function get_header()
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.ActiveSheet
'
'    Dim re As Object
'    Set re = CreateObject("VBScript.RegExp")
'    Dim s() As String
'    Dim r As Range
'    Dim rng As Range
'    Dim rs() As Integer
'    Dim arr As Variant: arr = Array(".*金融機関.*",".*支店.*",".*種別.*",".*口座名義.*",".*請求額.*",".*顧客コード.*")
'    ReDim s(Ubound(arr))
'    ReDim rs(Ubound(arr))
'    s = arr
'
'    Dim i As Integer
'    For i = Lbound(s) To Ubound(s)
'        With re
'            .Pattern = s(i)
'            .IgnoreCase = False
'            .Global = False
'            With ws
'                Set rng = Range( _
'                    .Range("A1"), _
'                    .Range("A1") _
'                        .SpecialCells(xlCellTypeVisible) _
'                            .SpecialCells(xlCellTypeLastCell) _
'                )
'            End With
'            For Each r In rng
'                If .test(r.Text) Then
'                    rs(i) = r.Row
'                    If i > 0 Then
'                        If rs(i) = rs(i-1) Then
'                            Exit For
'                        End If
'                    End If
'                    check_last = True
'                End If
'            Next r
'        End With
'    Next
'    Set rng = Nothing
'    Set re = Nothing
'    Set ws = Nothing
'End Function

Function check_indata(ByVal row As Integer, ByVal lastcol As Integer) As Boolean: check_indata = False
    '現在、顧客コード、請求額いずれかがない場合、True
    'リテラルで顧客コードと、請求額セル位置を判別。修正余地あり
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim rng As Range
    Set rng = Range(Cells(row, 1), Cells(row, lastcol))
    Dim r As Range
    For Each r In rng
        With r
            If .Text = "" Then
                Exit Function
            Else
                If .Column = 4 + 1 Or .Column = 5 + 1 Then
                    If Not IsNumeric(.Text) Then
                        Exit Function
                    End If
                Else
                    If IsNumeric(.Text) Then
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
    check_indata = True
    Set rng = Nothing
End Function

Private Sub get_indata_list()
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.ActiveSheet
'    With ws.Range("A1")
'        With .SpecialCells(xlCellTypeVisible)
'            Dim rngs As Range
'            Set rngs = .SpecialCells(xlCellTypeLastCell)
'            Dim rng As Range
'            For Each rng In rngs
'                Debug.Print rng.Text
'            Next rng
'        End With
'    End With
End Sub

Function get_indata(ByVal row As Integer, ByVal lastcol As Integer) As String()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    Dim s() As String
    ReDim s(lastcol)
    
    Dim i As Integer: i = 0
    Do
        s(i) = ws.Cells(row, i + 1).Text
        i = i + 1
    Loop Until i = lastcol
    
    get_indata = s
End Function

Function get_last_cell() As Long()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    Dim row As Integer
    Dim col As String

    Dim l(2) As Long
    
    With ws.Range("A1")
        With .SpecialCells(xlCellTypeVisible).SpecialCells(xlCellTypeLastCell)
            l(0) = .row
            '100件以上ない想定
            If .row > 100 Then
                l(0) = 100
            End If
            'l(1) = .Column
            '事故るので固定で7とした
            l(1) = 7
        End With
        If Trim(.Text) = "" Then
            l(2) = 1
        Else
            l(2) = CLng(.Text) + 1
        End If
    End With

    get_last_cell = l
End Function

Private Function get_grandparent_folder() As String
'   HISTORY
'      2020/03/14: 新規作成
    Dim fso As Object
    Dim tmp As String
    Set fso = CreateObject(FSOBJ)
    tmp = ThisWorkbook.Path
    tmp = Replace(tmp, fso.GetFolder(ThisWorkbook.Path).Name, "")
    tmp = Left(tmp, (Len(tmp) - 1))
    get_grandparent_folder = tmp
    Set fso = Nothing
End Function


Private Sub CreateInputTextData(ByRef v() As Variant)
    ' インプット設定
    '-------------------------------------------------------------------------'
    Dim inf As String: inf = v(0)
    Dim ibook As Workbook
    Dim isheet As Worksheet
    Set ibook = Workbooks.Open(inf, 0)
    ibook.Application.EnableEvents = False
    ibook.Application.DisplayAlerts = False
    ibook.Application.ScreenUpdating = False
    Set isheet = ibook.ActiveSheet
    '-------------------------------------------------------------------------'

    ' データ位置の取得
    '-------------------------------------------------------------------------'
    Dim last As Integer
    Dim first As Integer
    ' データが入っている開始行
    first = 1
    'Do Until Not IsNumeric(isheet.Cells(first, 1).Value)
    '    If first > 100 Then
    '        Exit Do
    '    End If
    '    first = first + 1
    'Loop

    ' データが入っている最終行
    last = isheet.Cells(isheet.Rows.Count, 1).End(xlUp).row
    Do Until IsNumeric(isheet.Cells(last, 1).Value)
        If last < 1 Then
            Exit Do
        End If
        last = last - 1
    Loop
    '-------------------------------------------------------------------------'

    ' データの取得・テキストデータ作成
    '-------------------------------------------------------------------------'
    Dim otf As String: otf = v(1)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object: Set ts = fso.OpenTextFile(otf, 2, -2)
    Dim data As String: data = vbNullString
    Dim drange As Range
    Dim drow As Range
    Dim cols(8) As String
    Set drange = isheet.Range(Cells(first, 1), Cells(last, 7))

    For Each drow In drange.Rows
        ' 高速化のため、単純な配列にセットし、それを参照させる
        cols(1) = RTrim(drow.Columns(1).Text)                                ' 顧客コード
        cols(2) = RTrim(drow.Columns(2).Text)                                ' 登録名
        cols(3) = RTrim(drow.Columns(3).Text)                                ' 銀行名
        cols(4) = RTrim(drow.Columns(4).Text)                                ' 支店名
        cols(5) = RTrim(drow.Columns(5).Text)                                ' 口座種別
        cols(6) = RTrim(RpaLibrary.ConvertToOmoji(drow.Columns(6).Text))     ' 口座名義
        cols(7) = RTrim(Replace(drow.Columns(7).Text, ",", vbNullString))    ' 金額
        If cols(1) <> vbNullString And IsNumeric(cols(1)) And Len(cols(1)) > 7 Then
            data = data & ","
            data = data & ","
            data = data & ","
            data = data & "," & cols(1)
            data = data & ","
            data = data & ","
            data = data & ","
            data = data & ","
            data = data & "," & cols(3)
            data = data & "," & cols(4)
            data = data & "," & cols(5)
            data = data & ","
            data = data & "," & cols(6)
            data = data & "," & cols(7)
            data = data & ","
            data = data & ","
            data = data & ","
            data = data & ","
            ts.WriteLine (data)
        End If
        Erase cols
        data = vbNullString
    Next
    ts.Close

    Set ts = Nothing
    Set fso = Nothing
    '-------------------------------------------------------------------------'

    ' 終了設定
    '-------------------------------------------------------------------------'
    Set json = Nothing
    Set isheet = Nothing
    ibook.Close
    Set ibook = Nothing
    '-------------------------------------------------------------------------'
End Sub

Private Sub list_caller()
'   NOTE
'       runtime_directory で実行
'
'   HISTORY
'      2020/02/03: 新フォーマット対応
'      2020/03/14: work.xlsm 対応
    Dim list As list_data
    Dim jsn As Object: Set jsn = GetProjectJson()
    'If ThisWorkbook.Name = WORK_XLSX Then
    '    list.book_name = _
    '        project_directory + "\" + "停止リスト.xls"
    '    list.output_csv_name = _
    '        ThisWorkbook.Path + "\" + "teishi_list_tmp.csv"
    'End If
    If ThisWorkbook.Name = RPA_XLM Then
        list.book_name = jsn("USR_PRJ_WRK") + "\" + "input.xls"
        list.output_csv_name = _
            jsn("USR_PRJ_WRK") + "\" + "teishi_list_tmp.csv"
    End If
    list.primary_data_column = 1
    Set list.book_obj = Workbooks.Open(list.book_name, 0)
    list.book_obj.Application.EnableEvents = False
    list.book_obj.Application.DisplayAlerts = False
    list.book_obj.Application.ScreenUpdating = False
    Set list.sheet_obj = list.book_obj.ActiveSheet
    Call get_list_last_row(list)
    Call list_main(list)
    'Application.EnableEvents = True
    'Application.DisplayAlerts = True
    list.book_obj.Close
    Set list.book_obj = Nothing
    Set list.sheet_obj = Nothing
    Set jsn = Nothing
End Sub

Private Sub get_list_last_row(ByRef list As list_data)
    list.last_data_row = list.sheet_obj.Cells(list.sheet_obj.Rows.Count, list.primary_data_column).End(xlUp).row
    Do Until IsNumeric(list.sheet_obj.Cells(list.last_data_row, list.primary_data_column))
        list.last_data_row = list.last_data_row - 1
        If list.last_data_row < 0 Then
            list.last_data_row = 0
            Exit Sub
        End If
    Loop
End Sub

Private Sub list_main(ByRef list As list_data)
    Dim i As Integer
    Dim fso As Object
    Set fso = CreateObject(FSOBJ)
    With fso
        Dim ts As Object
        Set ts = .OpenTextFile(list.output_csv_name, 2, -2)
        Dim code As String: code = vbNullString
        Dim txt As String: txt = vbNullString
        Dim dat As String: dat = vbNullString
        Dim rngs As Range
        Dim rng As Range
        With ts
            Set rngs = list.sheet_obj.Range(Cells(1, 1), Cells(list.last_data_row, 7))
            For Each rng In rngs.Rows
                Set list.kokyaku_code_cell = rng.Columns(list.primary_data_column)
                Set list.touroku_name_cell = rng.Columns(2)
                Set list.ginkou_name_cell = rng.Columns(3)
                Set list.siten_name_cell = rng.Columns(4)
                Set list.kouza_shubet_cell = rng.Columns(5)
                Set list.kouza_meigi_cell = rng.Columns(6)
                Set list.kingaku_cell = rng.Columns(7)

'-- 旧依頼書はこちらのレイアウト(廃止) ------------------------*
                'Set list.ginkou_name_cell = rng.Columns(1)
                'Set list.siten_name_cell = rng.Columns(2)
                'Set list.kouza_shubet_cell = rng.Columns(3)
                'Set list.kouza_meigi_cell = rng.Columns(4)
                'Set list.kingaku_cell = rng.Columns(5)
                'Set list.kokyaku_code_cell = rng.Columns(6)
                'Set list.touroku_name_cell = rng.Columns(7)
'--------------------------------------------------------------*

                If IsNull(list.ginkou_name_cell.Text) Then
                    list.ginkou_name = vbNullString
                Else
                    list.ginkou_name = list.ginkou_name_cell.Text
                End If
                If IsNull(list.siten_name_cell.Text) Then
                    list.siten_name = vbNullString
                Else
                    list.siten_name = list.siten_name_cell.Text
                End If
                If IsNull(list.kouza_shubet_cell.Text) Then
                    list.kouza_shubet = vbNullString
                Else
                    list.kouza_shubet = list.kouza_shubet_cell.Text
                End If
                If IsNull(list.kouza_meigi_cell.Text) Then
                    list.kouza_meigi = vbNullString
                Else
                    list.kouza_meigi = list.kouza_meigi_cell.Text
                    list.kouza_meigi = convert_to_omoji(list.kouza_meigi)
                End If
                If IsNull(list.kingaku_cell.Text) Then
                    list.kingaku = vbNullString
                Else
                    list.kingaku = list.kingaku_cell.Text
                    If InStr(list.kingaku, ",") > 0 Then
                        list.kingaku = Replace(list.kingaku, ",", "")
                    End If
                End If
                If IsNull(list.kokyaku_code_cell.Text) Then
                    list.kokyaku_code = vbNullString
                Else
                    list.kokyaku_code = list.kokyaku_code_cell.Text
                End If
                If IsNull(list.touroku_name_cell.Text) Then
                    list.touroku_name = vbNullString
                Else
                    list.touroku_name = list.touroku_name_cell.Text
                End If
                If Not list.kokyaku_code = vbNullString Then
                    If IsNumeric(list.kokyaku_code) Then
                        If Len(list.kokyaku_code) > 7 Then
                            txt = txt + ","
                            txt = txt + ","
                            txt = txt + ","
                            txt = txt + "," + list.kokyaku_code
                            txt = txt + ","
                            txt = txt + ","
                            txt = txt + ","
                            txt = txt + ","
                            txt = txt + "," + list.ginkou_name
                            txt = txt + "," + list.siten_name
                            txt = txt + "," + list.kouza_shubet
                            txt = txt + ","
                            txt = txt + "," + list.kouza_meigi
                            txt = txt + "," + list.kingaku
                            txt = txt + ","
                            txt = txt + ","
                            txt = txt + ","
                            txt = txt + ","
                            ts.WriteLine (txt)
                            txt = vbNullString
                        End If
                    End If
                End If
            Next
            .Close
        End With
        Set rng = Nothing
        Set ts = Nothing
    End With
    Set fso = Nothing
End Sub

Function get_startpos() As Long
    Dim tmp As String: tmp = ThisWorkbook.Path + "\position"
    Dim fso As Object
    Dim m As String
    Dim tmpf As Object
    Set fso = CreateObject(FSOBJ)
    With fso
        Set tmpf = .OpenTextFile(tmp, 1, True, -2)
        'ForReading, TristateUseDefault
        With tmpf
            If .AtEndOfStream Then
                get_startpos = 1
            Else
                get_startpos = CLng(Trim(.ReadLine)) + 1
            End If
            .Close
        End With
    End With
    Set fso = Nothing
    Set tmpf = Nothing
End Function

Sub set_start_position(ByVal cnt As Integer)
    Dim tmp As String: tmp = ThisWorkbook.Path + "\position"
    Dim fso As Object
    Dim m As String
    Dim tmpf As Object
    Set fso = CreateObject(FSOBJ)
    With fso
        Set tmpf = .OpenTextFile(tmp, 2, True, -2)
        'ForWriting
        'TristateUseDefault
        With tmpf
            .Write (Str(cnt))
            .Close
        End With
    End With
    Set fso = Nothing
    Set tmpf = Nothing
    'Range("A1") = cnt
End Sub

Private Sub get_shubet(ByRef iraisho As iraisho_data)
    Select Case iraisho.meisai_d.kouza_shubet
        Case "1"
            iraisho.meisai_d.kouza_shubet = "普通"
        Case "2"
            iraisho.meisai_d.kouza_shubet = "当座"
        Case Else
            iraisho.meisai_d.kouza_shubet = "種別エラー"
    End Select
End Sub

'Function trim_bankprf(v As Variant) As String
'    Dim s As String: s = v
'    If InStr(s, "銀行") > 0 Then
'        trim_bankprf = Replace(s, "銀行", "", vbTextCompare)
'    Else
'        trim_bankprf = s
'    End If
'End Function

'Function trim_bankprf(v As Variant) As String
'    Dim s As String: s = v
'    If InStr(s, "銀行") > 0 Then
'        trim_bankprf = Replace(s, "銀行", "", vbTextCompare)
'    Else
'        trim_bankprf = s
'    End If
'End Function

Private Function trim_banksfx(ByRef iraisho As iraisho_data, Optional indicate_sfx As String = vbNullString) As iraisho_data
    Dim sfx() As Variant: sfx = Array("銀行", "信用金庫", "信金")
    Dim i As Integer
    If indicate_sfx = vbNullString Then
        For i = LBound(sfx) To UBound(sfx)
            If InStr(iraisho.meisai_d.ginkou_name, sfx(i)) > 0 Then
                iraisho.meisai_d.ginkou_name = Replace(iraisho.meisai_d.ginkou_mei, sfx(i), "", vbTextCompare)
            End If
        Next
    Else
        If InStr(iraisho.meisai_d.ginkou_name, indicate_sfx) > 0 Then
            iraisho.meisai_d.ginkou_name = Replace(iraisho.meisai_d.ginkou_mei, indicate_sfx, "", vbTextCompare)
        End If
    End If
    trim_banksfx = iraisho
End Function

Function check_banksfx(ByVal s As String) As Boolean
    If InStr(s, "銀行") > 0 Then
        check_banksfx = "GINKO"
    ElseIf InStr(s, "信用金庫") > 0 Then
        check_banksfx = "SINKIN"
    ElseIf InStr(s, "信用組合") > 0 Then
        check_banksfx = "SINKUM"
    End If
End Function

Sub test()
    Debug.Print get_sis_target("福島信金庫")
End Sub

Private Function read_sheet_config(ByRef iraisho As iraisho_data, ByVal ky As String) As String: read_sheet_config = vbNullString
'   OVERVIEW
'      各種設定ファイルにある銀行コードに
'      依頼書データの銀行コードが一致していたら
'      その銀行コードに対応する銀行名を返す
'   HISTORY
'      2019/12/14 : 新規作成
'      2019/12/17 : マクロファイル対応                # 廃止
'      2020/08/13 : JSON対応, キーを銀行コードに変更

    Dim jsn As Object: Set jsn = GetProjectJson()
    Dim obj As Object: Set obj = jsn("ROO_PRJ_RPA")(ky)("TARGET_BANK")

    For Each bank_id In obj.Keys
        If iraisho.meisai_d.ginkou_code = bank_id Then
            read_sheet_config = bank_id
            Exit Function
        End If
    Next

    Set obj = Nothing
    Set jsn = Nothing

    ''Dim banks() As String: banks = get_list_of_file(iraisho.config_file_name)
    'Dim i As Integer
    'For i = LBound(banks) To UBound(banks)
    '    If iraisho.meisai_d.ginkou_name = banks(i) Then
    '        read_sheet_config = iraisho.meisai_d.ginkou_name
    '        Exit Function
    '    End If
    'Next
End Function


Private Function get_sis_target( _
    ByRef iraisho As iraisho_data) As String: get_sis_target = vbNullString
'   HISTORY
'      2019/10/25 対象エクセルファイルの外部化
'      2019/12/14 : 共通化対応
'      2019/12/17 : マクロファイル対応
'      2020/01/06 : ＳＩＳ対応
    Call get_ginkou_saffix(iraisho)
    If iraisho.meisai_d.ginkou_saffix = "信用金庫" Then
        get_sis_target = iraisho.meisai_d.ginkou_name
    End If
    'If ThisWorkbook.Name = WORK_XLSX Then
    '    iraisho.config_file_name = get_file_text(jsn("USR_PRJ_WRK") + "\runtime_dir") + "\sis"
    'ElseIf ThisWorkbook.Name = RPA_XLM Then
    '    iraisho.config_file_name = jsn("USR_PRJ_WRK") + "\sis"
    'End If
    'get_sis_target = read_sheet_config(iraisho)
End Function


Private Sub get_iraisho(ByRef iraisho As iraisho_data)
'   HISTORY
'      2019/12/12: 構造体対応
'      2019/12/12: ＳＩＳ依頼書対応
'      2020/08/13: JSON対応
'   NOTE
'      銀行名に対応する　ブック名
'      　　　　　　　　　シートレイアウト
'      　　　　　　　　　シート分け方
'      を指定
'   自営信金の依頼書と
'   個別の依頼書のレイアウトは同じNORMAL2
'
'
    Dim jsn As Object: Set jsn = GetProjectJson()
    Dim obj As Object: Set obj = Nothing
'
    Select Case iraisho.meisai_d.ginkou_code
        'Case get_gnet1_target(iraisho)
        Case read_sheet_config(iraisho, "GNET1")
            Set obj = jsn("ROO_PRJ_RPA")("GNET1")
            'iraisho.book_name = IRAISHO_GNET1
            'iraisho.sheet_layout = layout_type.GNET
            'iraisho.sheet_change_type = sheet_change.GINKOU
        'Case get_gnet2_target(iraisho)
        Case read_sheet_config(iraisho, "GNET2")
            Set obj = jsn("ROO_PRJ_RPA")("GNET2")
            'iraisho.book_name = IRAISHO_GNET2
            'iraisho.sheet_layout = layout_type.GNET
            'iraisho.sheet_change_type = sheet_change.GINKOU
        'Case get_gnet3_target(iraisho)
        Case read_sheet_config(iraisho, "GNET3")
            Set obj = jsn("ROO_PRJ_RPA")("GNET3")
            'iraisho.book_name = IRAISHO_GNET3
            'iraisho.sheet_layout = layout_type.GNET
            'iraisho.sheet_change_type = sheet_change.GINKOU
        'Case get_mizuho_target(iraisho)
        Case read_sheet_config(iraisho, "MIZUHO")
            Set obj = jsn("ROO_PRJ_RPA")("MIZUHO")
            'iraisho.book_name = IRAISHO_MIZUHO
            'iraisho.sheet_layout = layout_type.NORMAL
            'iraisho.sheet_change_type = sheet_change.SITEN
        'Case get_yucho_target(iraisho)
        Case read_sheet_config(iraisho, "YUCHO")
            Set obj = jsn("ROO_PRJ_RPA")("YUCHO")
            'iraisho.book_name = IRAISHO_YUCHO
            'iraisho.sheet_layout = layout_type.YUCHO
            'iraisho.sheet_change_type = sheet_change.GINKOU
        'Case get_mitsui_target(iraisho)
        Case read_sheet_config(iraisho, "MITSUI")
            Set obj = jsn("ROO_PRJ_RPA")("MITSUI")
            'iraisho.book_name = IRAISHO_MITSUI
            'iraisho.sheet_layout = layout_type.NORMAL2
            'iraisho.sheet_change_type = sheet_change.SITEN
        'Case get_taiko_target(iraisho)
        Case read_sheet_config(iraisho, "TAIKO")
            Set obj = jsn("ROO_PRJ_RPA")("TAIKO")
            'iraisho.book_name = IRAISHO_TAIKO
            'iraisho.sheet_layout = layout_type.NORMAL
            'iraisho.sheet_change_type = sheet_change.SITEN
        'Case get_joyo_target(iraisho)
        Case read_sheet_config(iraisho, "JOYO")
            Set obj = jsn("ROO_PRJ_RPA")("JOYO")
            'iraisho.book_name = IRAISHO_JOYO
            'iraisho.sheet_layout = layout_type.NORMAL2
            'iraisho.sheet_change_type = sheet_change.SITEN
        'Case get_akita_target(iraisho)
        Case read_sheet_config(iraisho, "AKITA")
            Set obj = jsn("ROO_PRJ_RPA")("AKITA")
            'iraisho.book_name = IRAISHO_AKITA
            'iraisho.sheet_layout = layout_type.NORMAL
            'iraisho.sheet_change_type = sheet_change.SITEN
        'Case get_asikaga_target(iraisho)
        Case read_sheet_config(iraisho, "ASIKAGA")
            Set obj = jsn("ROO_PRJ_RPA")("ASIKAGA")
            'iraisho.book_name = IRAISHO_ASIKAGA
            'iraisho.sheet_layout = layout_type.ASIKAGA
            'iraisho.sheet_change_type = sheet_change.GINKOU
        Case read_sheet_config(iraisho, "SIS_ZIEI")
            Set obj = jsn("ROO_PRJ_RPA")("SIS_ZIEI")
            '----------------------------------------------------'
            iraisho.meisai_d.ziei_code = obj("TARGET_BANK")(iraisho.meisai_d.ginkou_code)("Z-ID")
            '----------------------------------------------------'
            'iraisho.book_name = IRAISHO_SIS_ZIEI
            'iraisho.sheet_layout = layout_type.NORMAL2
            'iraisho.sheet_change_type = sheet_change.GINKOU
        Case read_sheet_config(iraisho, "SIS")
            Set obj = jsn("ROO_PRJ_RPA")("SIS")
            'iraisho.book_name = IRAISHO_SIS
            'iraisho.sheet_layout = layout_type.NORMAL
            'iraisho.sheet_change_type = sheet_change.GINKOU
        'Case get_sis_target(iraisho)
        '    Set obj = Nothing
        '    iraisho.book_name = IRAISHO_SIS
        '    iraisho.sheet_layout = layout_type.NORMAL
        '    iraisho.sheet_change_type = sheet_change.GINKOU
        Case read_sheet_config(iraisho, "NORMAL")
            Set obj = jsn("ROO_PRJ_RPA")("NORMAL")
            'iraisho.book_name = IRAISHO_NORMAL
            'iraisho.sheet_layout = layout_type.NORMAL
            'iraisho.sheet_change_type = sheet_change.GINKOU
        Case Else
            If iraisho.meisai_d.ginkou_name = get_sis_target(iraisho) Then
                'SIS --------------------------------------------------------------------'
                Set obj = Nothing
                iraisho.book_name = IRAISHO_SIS
                iraisho.sheet_layout = layout_type.NORMAL
                iraisho.sheet_change_type = sheet_change.GINKOU
                '------------------------------------------------------------------------'
            Else
                'NORMAL -----------------------------------------------------------------'
                Set obj = Nothing
                iraisho.book_name = IRAISHO_NORMAL
                iraisho.sheet_layout = layout_type.NORMAL
                iraisho.sheet_change_type = sheet_change.GINKOU
                '------------------------------------------------------------------------'
            End If
    End Select


    If Not obj Is Nothing Then
        iraisho.book_name = obj("Book_name")
        iraisho.meisai_d.ginkou_name = obj("TARGET_BANK")(iraisho.meisai_d.ginkou_code)("Name")
        iraisho.sheet_layout = obj("Layout_type")
        iraisho.sheet_change_type = obj("Resheet_type")
    End If

    Set obj = Nothing
    Set jsn = Nothing
End Sub

Sub testaba()
    Dim wb As Workbook
    Set wb = get_iraisho("あかぎ信用組合")
End Sub

'Private Sub get_project_config(ByRef project As rpa_project)
'    Select Case ThisWorkbook.Name
'        Case RPA_XLM
'            'Dim fso As Object
'            'project.json_name = ThisWorkbook.Path + "\" + "rpa_context.json"
''
'            'Dim jobj As Object
'            'Set fso = CreateObject(FSOBJ)
'            'Dim ts As Object
'            'Set ts = fso.OpenTextFile(project.json_name)
'            'Set jobj = JsonConverter.ParseJson(ts.ReadAll())
'            'project.project_dir = jobj("project_dir")
'            'project.runtime_dir = jobj("runtime_dir")
'            'Set jobj = Nothing
'            'Set ts = Nothing
'            'Set fso = Nothing
'        Case WORK_XLSX
'            project.project_dir = _
'                get_file_text(ThisWorkbook.Path + "\" + "project_dir")
'            project.runtime_dir = _
'                get_file_text(ThisWorkbook.Path + "\" + "runtime_dir")
'    End Select
'End Sub

Private Sub iraisho_initializer(ByRef iraisho As iraisho_data)
'   HISTORY
'      2019/12/14 : 新規作成
    iraisho.sheet_layout = 0
    iraisho.sheet_change_type = 0
    iraisho.write_row = 0
    iraisho.max_write_row = 0

    iraisho.start_flag = False
    iraisho.start_code = vbNullString
    iraisho.writer_name = vbNullString
    iraisho.config_file_name = vbNullString
    iraisho.genpon_sheet_name = vbNullString
    iraisho.write_sheet_name = vbNullString
    iraisho.book_name = vbNullString
    iraisho.sheet_m = vbNullString
    Set iraisho.genpon_sheet_obj = Nothing
    Set iraisho.write_sheet_obj = Nothing
    Set iraisho.book_obj = Nothing
    Set iraisho.writer_name_cell = Nothing
    Set iraisho.ginkou_name_cell = Nothing
    Set iraisho.siten_name_cell = Nothing
    Set iraisho.kouza_shubet_cell = Nothing
    Set iraisho.kouza_num_cell = Nothing
    Set iraisho.kouza_meigi_cell = Nothing
    Set iraisho.kingaku_cell = Nothing
    Set iraisho.bikou_cell = Nothing
    Set iraisho.ginkou_code_cell = Nothing
    Set iraisho.siten_code_cell = Nothing
    Set iraisho.furikae_date_cell = Nothing
    Set iraisho.tucho_num_cell = Nothing
    Set iraisho.ginkou_code1_cell = Nothing
    Set iraisho.ginkou_code2_cell = Nothing
    Set iraisho.ginkou_code3_cell = Nothing
    Set iraisho.ginkou_code4_cell = Nothing
    Set iraisho.ginkou_saffix_cell = Nothing
    Set iraisho.kokyaku_code_cell = Nothing
    Set iraisho.funo_code_cell = Nothing
    iraisho.meisai_d.ginkou_name = vbNullString
    iraisho.meisai_d.siten_name = vbNullString
    iraisho.meisai_d.kouza_shubet = vbNullString
    iraisho.meisai_d.kouza_num = vbNullString
    iraisho.meisai_d.kouza_meigi = vbNullString
    iraisho.meisai_d.kingaku = vbNullString
    iraisho.meisai_d.bikou = vbNullString
    iraisho.meisai_d.ginkou_code = vbNullString
    iraisho.meisai_d.siten_code = vbNullString
    iraisho.meisai_d.furikae_date = vbNullString
    iraisho.meisai_d.tucho_num = vbNullString
    iraisho.meisai_d.ginkou_code1 = vbNullString
    iraisho.meisai_d.ginkou_code2 = vbNullString
    iraisho.meisai_d.ginkou_code3 = vbNullString
    iraisho.meisai_d.ginkou_code4 = vbNullString
    iraisho.meisai_d.ginkou_saffix = vbNullString
    iraisho.meisai_d.kokyaku_code = vbNullString
    iraisho.meisai_d.funo_code = vbNullString
End Sub

Private Sub get_iraisho_start_code(ByRef iraisho As iraisho_data)
'   HISTORY
'      2020/01/28: 新規作成
'      2020/03/14: work.xlsm対応
    Dim jsn As Object: Set jsn = GetProjectJson()
    iraisho.start_code = get_file_text(jsn("USR_PRJ_SYS") + "\" + "start_code")
    Set jsn = Nothing
End Sub

Private Sub set_iraisho_start_code( _
    ByRef meisai As meisai_data, _
      ByRef list As list_data)
'   HISTORY
'      2020/01/28: 新規作成
'      2020/03/14: work.xlsm対応
    Dim fso As Object
    Dim ts As Object
    Dim jsn As Object: Set jsn = GetProjectJson()

    Set fso = CreateObject(FSOBJ)
    Set ts = fso.OpenTextFile(jsn("USR_PRJ_SYS") + "\" + "start_code", 2, -2)

    list.book_name = jsn("USR_PRJ_WRK") + "\" + "input.xls"

    Set list.book_obj = Workbooks.Open(list.book_name, 0)
    Set list.sheet_obj = list.book_obj.ActiveSheet

    If list.sheet_obj.Cells(1, 1).Text = "END" Then
        ts.WriteLine ("XXXXXXXXXXXXX")
    Else
        ts.WriteLine (meisai.kokyaku_code)
    End If
    ts.Close
    Set ts = Nothing

    list.book_obj.Close
    Set list.book_obj = Nothing
    Set list.sheet_obj = Nothing

    Set fso = Nothing
    Set jsn = Nothing
End Sub

Private Sub set_meisai_data(ByRef meisai As meisai_data, ByRef data() As String)
'   HISTORY
'      2020/01/28: 新規作成
'      data の構造化対応は要件等
    meisai.ginkou_name = data(9)
    meisai.ginkou_name = RTrim(meisai.ginkou_name)
    meisai.siten_name = data(10)
    meisai.siten_name = RTrim(meisai.siten_name)
    meisai.kouza_shubet = data(11)
    meisai.kouza_shubet = RTrim(meisai.kouza_shubet)
    meisai.kouza_num = data(12)
    meisai.kouza_num = RTrim(meisai.kouza_num)
    meisai.kouza_meigi = data(13)
    meisai.kouza_meigi = RTrim(meisai.kouza_meigi)
    meisai.kingaku = data(14)
    meisai.kingaku = RTrim(meisai.kingaku)
    meisai.bikou = "株式会社　モテキ"
'   meisai.bikou = RTrim(meisai.bikou)
    meisai.ginkou_code = data(7)
    meisai.ginkou_code = RTrim(meisai.ginkou_code)
    meisai.siten_code = data(8)
    meisai.siten_code = RTrim(meisai.siten_code)
    meisai.furikae_date = data(15)
    meisai.furikae_date = RTrim(meisai.furikae_date)
End Sub

Private Sub iraisho_caller()
'   HISTORY
'      2019/12/12: 構造体対応
'      2019/12/26: 足利銀行追加
'      2020/01/09: EOL判定のバグを修正
'      2020/01/28: 途中開始する対応
'                  set_iraisho_start_code
'                  get_iraisho_start_code
'      2020/02/01: 件数集計ロジックを追加
'      2020/03/14: work対応
'      2020/03/27: jsn対応

    Dim fso As Object
    Dim jsn As Object
    Dim meisai As meisai_data
    Dim iraisho As iraisho_data
    Dim list As list_data
'-- 件数集計ロジック ------------------*
    Dim summary() As summary_data
    ReDim summary(0)
    Dim i As Integer
    Dim hit As Boolean
    Dim total_count As Integer: total_count = 0
'--------------------------------------*
'
    Call iraisho_initializer(iraisho)
    'Call get_project_config(iraisho.project)
    Call get_iraisho_start_code(iraisho)
'
    Set jsn = GetProjectJson()
    Set fso = CreateObject(FSOBJ)
    With fso
        Dim ts As Object
        'Set ts = fso.OpenTextFile( _
        '    get_file_text(ThisWorkbook.Path + "\runtime_dir") + "\" + "teishi_write.csv", 1, -2)
        Set ts = fso.OpenTextFile(jsn("USR_PRJ_WRK") + "\" + "teishi_write.csv", 1, -2)
        With ts
            Dim txt As String: txt = vbNullString
            Dim data() As String
            Dim work_flag As Boolean: work_flag = False
            Do
                txt = .ReadLine()
                data = Split(txt, ",")
'-- 件数集計ロジック-----------------------------------------------*
                total_count = total_count + 1
                i = 0
                hit = False
                meisai.ginkou_name = data(9)
                meisai.ginkou_name = RTrim(meisai.ginkou_name)
                For i = LBound(summary) To UBound(summary)
                    If meisai.ginkou_name = summary(i).ginkou_name Then
                        summary(i).data_count = summary(i).data_count + 1
                        hit = True
                    End If
                Next
                If Not hit Then
                    summary(UBound(summary)).ginkou_name = meisai.ginkou_name
                    summary(UBound(summary)).data_count = 1
                    ReDim Preserve summary(UBound(summary) + 1)
                End If
'------------------------------------------------------------------*
'-- 途中開始する対応 ----------------------------------------------*
                meisai.kokyaku_code = data(4)
                meisai.kokyaku_code = RTrim(meisai.kokyaku_code)
                If Not iraisho.start_flag Then
                    If work_flag Then
'                       2. このタイミングでiraisho.start_flag をTrue にする
                        iraisho.start_flag = True
                    End If
'
                    If iraisho.start_code = "XXXXXXXXXXXXX" Then
                        iraisho.start_flag = True
                    Else
                        If iraisho.start_code = meisai.kokyaku_code Then
'                           1. 一致したコードが前回最後のデータなので、この段階では
'                              iraisho.start_flag はTrue にしない
                            work_flag = True
                        End If
                    End If
                End If

                If iraisho.start_flag Then
                    Call set_meisai_data(meisai, data)
                    iraisho.meisai_d = meisai
                    Call iraisho_controller(iraisho)
                End If
'------------------------------------------------------------------*
            Loop Until .AtEndOfStream

            Call set_iraisho_start_code(meisai, list)
'-- 件数集計ロジック-----------------------------------------------*
            summary(UBound(summary)).ginkou_name = "合計"
            summary(UBound(summary)).data_count = total_count
            Call write_summary(summary)
'------------------------------------------------------------------*
            Call iraisho_initializer(iraisho)
        End With
        Set ts = Nothing
    End With
    Set fso = Nothing
    Set jsn = Nothing
End Sub

Private Sub write_summary(ByRef summary() As summary_data)
'-- NOTE
'      runtime_directory で実行(iraisho caller から呼び出し)
'-- HISTORY
'   2020/02/01: 新規作成
    Dim fso As Object
    Dim i As Integer
    Dim jsn As Object: Set jsn = GetProjectJson()
    Set fso = CreateObject(FSOBJ)
    With fso
        Dim ts As Object
        Set ts = fso.OpenTextFile(jsn("USR_PRJ_WRK") + "\" + "summary.csv", 2, -2)
        With ts
            .WriteLine ("銀行名,件数")
            For i = LBound(summary) To UBound(summary)
                .WriteLine (summary(i).ginkou_name + "," + summary(i).data_count)
            Next
            .Close
        End With
        Set ts = Nothing
    End With
    Set fso = Nothing
    Set jsn = Nothing
End Sub

Private Sub iraisho_controller(ByRef iraisho As iraisho_data)
'-- HISTORY
'   XXXX/XX/XX: 氏名の大文字化
'   2019/12/12: 構造体対応
'
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim jsn As Object: Set jsn = GetProjectJson()
    
    '--- GET BOOK OBJECT -----------------------------------------
    Call get_iraisho(iraisho)
    Set iraisho.book_obj = Workbooks.Open( _
         jsn("USR_PRJ_WRK") + "\" + iraisho.book_name, 0)
    iraisho.book_obj.Application.EnableEvents = False
    iraisho.book_obj.Application.DisplayAlerts = False
    iraisho.book_obj.Application.ScreenUpdating = False
    '-------------------------------------------------------------

    iraisho.book_obj.Activate
    Call iraisho_main(iraisho)
    iraisho.book_obj.Save
    iraisho.book_obj.Close
    Set iraisho.book_obj = Nothing

    'Application.EnableEvents = True
    'Application.DisplayAlerts = True
    Set jsn = Nothing
End Sub


Private Function pad_left_zero(ByRef s As String, ByRef b As Integer) As String: pad_left_zero = s
    Dim l As Integer: l = b - Len(s)
    If l >= 0 Then
        pad_left_zero = String(l, "0") + s
    End If
End Function

Sub testxx()
    Dim s(10) As String
    s(0) = "東和銀行"
    s(1) = "hoge"
    s(2) = "1"
    s(3) = "０３３４４９８"
    s(4) = "ﾊﾔｼ"
    s(5) = "5389"
    s(6) = "モテキ（株）"
    s(7) = "221"
    s(8) = "6666"
    s(9) = "2019/10/15"
    Call out_main(s)
End Sub

Private Sub CreateIraisho(ByRef v() As Variant)
    Dim inf As String: inf = v(0)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object: Set ts = fso.OpenTextFile(inf, 1, -2)
    Dim wbook As Workbook: Set wbook = Nothing
    Dim wsheet As Worksheet: Set wsheet = Nothing
    Dim ws As Worksheet
    Dim bookname As String: bookname = vbNullString
    Dim line As String
    Dim vs() As String
    Dim idx As Integer

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Do Until ts.AtEndOfStream
        line = ts.ReadLine()
        vs = Split(line, ",")

        ' ブックの変更
        If bookname <> vs(0) Then
            If Not wbook Is Nothing Then
                wbook.Save
                wbook.Application.EnableEvents = True
                wbook.Application.DisplayAlerts = True
                wbook.Application.ScreenUpdating = True
                wbook.Close
                Set wbook = Nothing
            End If

            bookname = vs(0)
            Set wbook = Workbooks.Open(bookname, 0)
            wbook.Application.EnableEvents = False
            wbook.Application.DisplayAlerts = False
            wbook.Application.ScreenUpdating = False
        End If

        ' シートの変更
        If sheetname <> vs(2) Then
            If Not wsheet Is Nothing Then
                Set wsheet = Nothing
            End If

            sheetname = vs(2)
            If RpaLibrary.IsWorksheetExists(wbook, sheetname) Then
                Set wsheet = wbook.Worksheets(sheetname)
            Else
                wbook.Worksheets(vs(1)).Copy after:=Worksheets(vs(1))
                Set wsheet = wbook.ActiveSheet
                wsheet.Name = sheetname
            End If
        End If

        For idx = 4 To 116 Step 2
            If vs(idx) <> vbNullString Then
                wsheet.Range(vs(idx)).Value = vs(idx - 1)
            End If
        Next
    Loop

    ' 最後のシート
    If Not wsheet Is Nothing Then
        Set wsheet = Nothing
    End If

    ' 最後のブック
    If Not wbook Is Nothing Then
        wbook.Save
        wbook.Application.EnableEvents = True
        wbook.Application.DisplayAlerts = True
        wbook.Application.ScreenUpdating = True
        wbook.Close
        Set wbook = Nothing
    End If

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

'Sub test()
'    Dim s(8) As String
'    s(0) = "東和銀行"
'    s(1) = "hoge"
'    s(2) = "1"
'    s(3) = "０３３４４９８"
'    s(4) = "ﾊﾔｼ"
'    s(5) = "5389"
'    s(6) = "モテキ（株）"
'    s(7) = "221"
'    s(8) = "6666"
'    Call print_yuucho(s, ThisWorkbook)
'End Sub

Private Sub trim_ginkou_saffix(ByRef iraisho As iraisho_data)
'   HISTORY
'      - 2019/12/14 : 新規作成
    If InStr(iraisho.meisai_d.ginkou_name, "銀行") > 0 Then
        iraisho.meisai_d.ginkou_name = Replace(iraisho.meisai_d.ginkou_name, "銀行", "")
        iraisho.meisai_d.ginkou_saffix = "銀行"
    End If
    If InStr(iraisho.meisai_d.ginkou_name, "信用金庫") > 0 Then
        iraisho.meisai_d.ginkou_name = Replace(iraisho.meisai_d.ginkou_name, "信用金庫", "")
        iraisho.meisai_d.ginkou_saffix = "信用金庫"
    End If
    '--- 信金が正式名称ということはあるのか？----------------------------'
    If InStr(iraisho.meisai_d.ginkou_name, "信金") > 0 Then
        iraisho.meisai_d.ginkou_name = Replace(iraisho.meisai_d.ginkou_name, "信金", "")
        iraisho.meisai_d.ginkou_saffix = "信用金庫"
    End If
    '--------------------------------------------------------------------'
End Sub

Private Sub get_ginkou_saffix(ByRef iraisho As iraisho_data)
'   HISTORY
'      - 2020/01/06 : 新規作成
    If InStr(iraisho.meisai_d.ginkou_name, "銀行") > 0 Then
        iraisho.meisai_d.ginkou_saffix = "銀行"
    End If
    If InStr(iraisho.meisai_d.ginkou_name, "信用金庫") > 0 Then
        iraisho.meisai_d.ginkou_saffix = "信用金庫"
    End If
    '--- 信金が正式名称ということはあるのか？----------------------------'
    If InStr(iraisho.meisai_d.ginkou_name, "信金") > 0 Then
        iraisho.meisai_d.ginkou_saffix = "信用金庫"
    End If
    '--------------------------------------------------------------------'
End Sub

Private Sub get_tucho_num(ByRef iraisho As iraisho_data)
'   ゆうちょ専用。ロジック見栄えの為、サブルーチン化
'   HISTORY
'      - 2019/12/14 : 新規作成
    iraisho.meisai_d.kouza_num = Right(iraisho.meisai_d.kouza_num, 11)
    Dim no As String: no = iraisho.meisai_d.kouza_num
    no = no + "1"
    iraisho.meisai_d.kouza_num = Right(no, 8)
    no = Replace(no, iraisho.meisai_d.kouza_num, "")
    If Len(no) < 4 Then
        no = String(4 - Len(no), "0") + no
    End If
    no = "1" + no
    iraisho.meisai_d.tucho_num = no
End Sub

Private Sub iraisho_main(ByRef iraisho As iraisho_data)
    '2020-07-08 : 足利依頼書の備考欄の文字列を変更
    '             「株式会社モテキ」 -> 「(株) モテキ」
    '2020/08/13 : 自営信金コードの入力追加
    ' This Procedure is not EntryPoint
    '--- DATA SET
    'iraisho.meisai_d.ginkou_name = RTrim(iraisho.meisai_d.ginkou_name)
    'iraisho.meisai_d.siten_name   = meisai.siten_m
    'iraisho.meisai_d.kouza_shubet = meisai.shubetu
    iraisho.meisai_d.kouza_meigi = StrConv(iraisho.meisai_d.kouza_meigi, vbWide)
    'iraisho.meisai_d.kingaku      = iraisho.meisai_d.kingaku
    'iraisho.meisai_d.bikou        = iraisho.meisai_d.bikou
    'iraisho.meisai_d.ginkou_code  = iraisho.meisai_d.ginko_c
    'iraisho.meisai_d.siten_code   = iraisho.meisai_d.siten_c
    iraisho.meisai_d.furikae_date = Trim(Str(Int(Day(iraisho.meisai_d.furikae_date))))
    iraisho.writer_name = "林"

    Select Case iraisho.sheet_layout
        Case layout_type.GNET
            '--- DATA SET
            If iraisho.book_name <> IRAISHO_GNET3 Then
                iraisho.genpon_sheet_name = "原本(" + iraisho.meisai_d.ginkou_name + ")"
            Else
                iraisho.genpon_sheet_name = "原本"
            End If
            iraisho.write_sheet_name = iraisho.meisai_d.ginkou_name
            iraisho.max_write_row = 37
            iraisho.primary_data_column = 15
            iraisho.meisai_d.ginkou_code1 = Left(iraisho.meisai_d.ginkou_code, 1)
            iraisho.meisai_d.ginkou_code2 = Mid(iraisho.meisai_d.ginkou_code, 2, 1)
            iraisho.meisai_d.ginkou_code3 = Mid(iraisho.meisai_d.ginkou_code, 3, 1)
            iraisho.meisai_d.ginkou_code4 = Right(iraisho.meisai_d.ginkou_code, 1)
            iraisho.meisai_d.siten_code = StrConv( _
                                               pad_left_zero( _
                                                   iraisho.meisai_d.siten_code, 3), vbWide)
            iraisho.meisai_d.kouza_num = Right(iraisho.meisai_d.kouza_num, 7)
            iraisho.meisai_d.kouza_num = StrConv( _
                                               pad_left_zero( _
                                                  iraisho.meisai_d.kouza_num, 7), vbWide)
            '
            '--- SHEET SELECT
            Call change_sheet(iraisho)
            Call get_write_row(iraisho)
            If iraisho.write_row = iraisho.max_write_row Then
                iraisho.write_sheet_name = iraisho.meisai_d.ginkou_name + "I"
                Call change_sheet(iraisho)
                Call get_write_row(iraisho)
            End If
            '
            '--- SET CELL OBJ
            Set iraisho.ginkou_code1_cell = Cells(25, 18)
            Set iraisho.ginkou_code2_cell = Cells(25, 20)
            Set iraisho.ginkou_code3_cell = Cells(25, 22)
            Set iraisho.ginkou_code4_cell = Cells(25, 24)
            Set iraisho.furikae_date_cell = Cells(24, 26)
            Set iraisho.siten_code_cell = Cells(iraisho.write_row + 1, 2)
            Set iraisho.kouza_shubet_cell = Cells(iraisho.write_row + 1, 5)
            Set iraisho.kouza_num_cell = Cells(iraisho.write_row + 1, 8)
            Set iraisho.kingaku_cell _
                = Cells(iraisho.write_row + 1, iraisho.primary_data_column)
            '                                  ---------------------------
            Set iraisho.kouza_meigi_cell = Cells(iraisho.write_row + 1, 23)
            'Set iraisho.writer_name_cell  = Cells(xx, xx)
            '
            '--- DATA WRITE
            iraisho.ginkou_code1_cell.Value = iraisho.meisai_d.ginkou_code1
            iraisho.ginkou_code2_cell.Value = iraisho.meisai_d.ginkou_code2
            iraisho.ginkou_code3_cell.Value = iraisho.meisai_d.ginkou_code3
            iraisho.ginkou_code4_cell.Value = iraisho.meisai_d.ginkou_code4
            iraisho.furikae_date_cell.Value = iraisho.meisai_d.furikae_date
            iraisho.siten_code_cell.Value = iraisho.meisai_d.siten_code
            iraisho.kouza_shubet_cell.Value = iraisho.meisai_d.kouza_shubet
            iraisho.kouza_num_cell.Value = iraisho.meisai_d.kouza_num
            iraisho.kingaku_cell.Value = iraisho.meisai_d.kingaku
            iraisho.kouza_meigi_cell.Value = iraisho.meisai_d.kouza_meigi
            'iraisho.writer_name_cell.Value  = iraisho.writer_name
            '
        Case layout_type.NORMAL, layout_type.NORMAL2
            'HISTORY
            ' - 2020/01/08 : 原本シート名は"原本"に統一
            ' - 委託者コードのあるなしの違いだけなので、統一したい
            '
            '--- DATA SET
            iraisho.genpon_sheet_name = "原本"
            'Select Case iraisho.sheet_layout
            '    Case layout_type.NORMAL
            '        iraisho.genpon_sheet_name = "原本"
            '    Case layout_type.NORMAL2
            '        iraisho.genpon_sheet_name = "原本(自営)"
            'End Select
            Select Case iraisho.sheet_change_type
                Case sheet_change.GINKOU
                    iraisho.write_sheet_name = iraisho.meisai_d.ginkou_name
                Case sheet_change.SITEN
                    iraisho.write_sheet_name = iraisho.meisai_d.siten_name
            End Select
            iraisho.max_write_row = 23
            iraisho.primary_data_column = 16
            Call trim_ginkou_saffix(iraisho)
            '    => iraisho.meisai_d.ginkou_name  = bank name removed saffix
            '    => iraisho.meisai_d.ginkou_saffix   = bank saffix
            iraisho.meisai_d.kouza_num = Right(iraisho.meisai_d.kouza_num, 7)
            Call get_shubet(iraisho)
            '    => iraisho.meisai_d.kouza_shubet = kouza shubet character
            '
            '--- SHEET SELECT
            Call change_sheet(iraisho)
            Call get_write_row(iraisho)
            If iraisho.write_row = iraisho.max_write_row Then
                iraisho.write_sheet_name = iraisho.meisai_d.ginkou_name + "I"
                Call change_sheet(iraisho)
                Call get_write_row(iraisho)
            End If
            '
            '--- SET CELL OBJ
            Set iraisho.furikae_date_cell = Cells(3, 14)
            Set iraisho.ginkou_saffix_cell = Cells(4, 3)
            Set iraisho.ginkou_name_cell = Cells(4, 2)
            Set iraisho.siten_name_cell = Cells(iraisho.write_row + 1, 2)
            Set iraisho.kouza_shubet_cell = Cells(iraisho.write_row + 1, 3)
            Set iraisho.kouza_num_cell = Cells(iraisho.write_row + 1, 4)
            Set iraisho.kouza_meigi_cell = Cells(iraisho.write_row + 1, 6)
            Set iraisho.kingaku_cell _
                = Cells(iraisho.write_row + 1, iraisho.primary_data_column)
            '                                  ---------------------------
            Set iraisho.bikou_cell = Cells(iraisho.write_row + 1, 18)
            Select Case iraisho.sheet_layout
                Case layout_type.NORMAL
                    Set iraisho.writer_name_cell = Cells(8, 17)
                Case layout_type.NORMAL2
                    Set iraisho.writer_name_cell = Cells(9, 17)
            End Select

            '
            '--- DATA WRITE
            iraisho.ginkou_saffix_cell.Value = iraisho.meisai_d.ginkou_saffix
            iraisho.ginkou_name_cell.Value = iraisho.meisai_d.ginkou_name
            iraisho.furikae_date_cell.Value = iraisho.meisai_d.furikae_date
            iraisho.siten_name_cell.Value = iraisho.meisai_d.siten_name
            iraisho.kouza_shubet_cell.Value = iraisho.meisai_d.kouza_shubet
            iraisho.kouza_num_cell.Value = iraisho.meisai_d.kouza_num
            iraisho.kouza_meigi_cell.Value = iraisho.meisai_d.kouza_meigi
            iraisho.kingaku_cell.Value = iraisho.meisai_d.kingaku
            iraisho.bikou_cell.Value = iraisho.meisai_d.bikou
            iraisho.writer_name_cell.Value = iraisho.writer_name

            '--- 自営信金コードのセット
            If iraisho.book_name = IRAISHO_SIS_ZIEI Then
                Set iraisho.ziei_code_cell = Cells(8, 17)
                iraisho.ziei_code_cell.Value = iraisho.meisai_d.ziei_code
            End If

        Case layout_type.YUCHO
            '--- DATA SET
            iraisho.genpon_sheet_name = "原本"
            iraisho.write_sheet_name = "I"
            iraisho.max_write_row = 38
            iraisho.primary_data_column = 9
            Call get_shubet(iraisho)
            '    => iraisho.meisai_d.kouza_shubet = kouza shubet character
            Call trim_ginkou_saffix(iraisho)
            '    => iraisho.meisai_d.ginkou_name  = bank name removed saffix
            '    => iraisho.meisai_d.ginkou_saffix   = bank saffix
            Call get_tucho_num(iraisho)
            '    => iraisho.meisai_d.tucho_num    = yucho tucho number
            iraisho.meisai_d.tucho_num = _
                StrConv(iraisho.meisai_d.tucho_num, vbWide)
            iraisho.meisai_d.kouza_num = _
                StrConv(iraisho.meisai_d.kouza_num, vbWide)
            '
            '--- SHEET SELECT
            Do
                'Loop外でidxを指定してしまうとparentが固定になってしまう
                Call change_sheet(iraisho)
                Call get_write_row(iraisho)
                If iraisho.write_row < iraisho.max_write_row Then
                    Exit Do
                Else
                    iraisho.write_sheet_name = iraisho.write_sheet_name + "I"
                End If
            Loop
            '
            ' SET CELL OBJ
            Set iraisho.furikae_date_cell = Cells(22, 7)
            Set iraisho.tucho_num_cell = Cells(iraisho.write_row + 1, 1)
            Set iraisho.kouza_num_cell = Cells(iraisho.write_row + 1, 3)
            Set iraisho.kouza_meigi_cell = Cells(iraisho.write_row + 1, 5)
            Set iraisho.kingaku_cell _
                = Cells(iraisho.write_row + 1, iraisho.primary_data_column)
            '                                  ---------------------------
            Set iraisho.bikou_cell = Cells(iraisho.write_row + 1, 10)
            Set iraisho.writer_name_cell = Cells(12, 9)
            '
            ' DATA WRITE
            iraisho.furikae_date_cell.Value = iraisho.meisai_d.furikae_date
            iraisho.tucho_num_cell.Value = iraisho.meisai_d.tucho_num
            iraisho.kouza_num_cell.Value = iraisho.meisai_d.kouza_num
            iraisho.kouza_meigi_cell.Value = iraisho.meisai_d.kouza_meigi
            iraisho.kingaku_cell.Value = iraisho.meisai_d.kingaku
            iraisho.bikou_cell.Value = iraisho.meisai_d.bikou
            iraisho.writer_name_cell.Value = iraisho.writer_name
        Case layout_type.ASIKAGA
            '--- DATA SET
            iraisho.genpon_sheet_name = "原本"
            iraisho.write_sheet_name = "I"
            iraisho.max_write_row = 25
            iraisho.primary_data_column = 16
            iraisho.meisai_d.funo_code = "8"
            Call get_shubet(iraisho)
            '    => iraisho.meisai_d.kouza_shubet = kouza shubet character
            '
            '--- SHEET SELECT
            Do
                'Loop外でidxを指定してしまうとparentが固定になってしまう
                Call change_sheet(iraisho)
                Call get_write_row(iraisho)
                If iraisho.write_row < iraisho.max_write_row Then
                    Exit Do
                Else
                    iraisho.write_sheet_name = iraisho.write_sheet_name + "I"
                End If
            Loop
            '
            ' SET CELL OBJ
            Set iraisho.furikae_date_cell = Cells(3, 12)
            Set iraisho.siten_name_cell = Cells(iraisho.write_row + 1, 2)
            Set iraisho.siten_code_cell = Cells(iraisho.write_row + 1, 3)
            Set iraisho.kouza_shubet_cell = Cells(iraisho.write_row + 1, 4)
            Set iraisho.kouza_num_cell = Cells(iraisho.write_row + 1, 5)
            Set iraisho.kokyaku_code_cell = Cells(iraisho.write_row + 1, 6)
            Set iraisho.kouza_meigi_cell = Cells(iraisho.write_row + 1, 10)
            Set iraisho.kingaku_cell _
                = Cells(iraisho.write_row + 1, iraisho.primary_data_column)
            '                                  ---------------------------
            Set iraisho.funo_code_cell = Cells(iraisho.write_row + 1, 18)
            Set iraisho.bikou_cell = Cells(iraisho.write_row + 1, 19)
            '
            ' DATA WRITE
            iraisho.furikae_date_cell.Value = iraisho.meisai_d.furikae_date
            iraisho.siten_name_cell.Value = iraisho.meisai_d.siten_name
            iraisho.siten_code_cell.Value = iraisho.meisai_d.siten_code
            iraisho.kouza_shubet_cell.Value = iraisho.meisai_d.kouza_shubet
            iraisho.kouza_num_cell.Value = iraisho.meisai_d.kouza_num
            iraisho.kokyaku_code_cell.Value = iraisho.meisai_d.kokyaku_code
            iraisho.kouza_meigi_cell.Value = iraisho.meisai_d.kouza_meigi
            iraisho.kingaku_cell.Value = iraisho.meisai_d.kingaku
            iraisho.funo_code_cell.Value = iraisho.meisai_d.funo_code

            '2020-07-18: 対応
            'iraisho.bikou_cell.Value = iraisho.meisai_d.bikou
            iraisho.bikou_cell.Value = "(株) モテキ"
        Case Else
    End Select
End Sub

Private Sub get_write_row(ByRef iraisho As iraisho_data)
    iraisho.write_row = Cells(100, iraisho.primary_data_column).End(xlUp).row
End Sub

Private Sub change_sheet(ByRef iraisho As iraisho_data)
    Dim i As Integer
    Set iraisho.write_sheet_obj = Nothing
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If Worksheets(i).Name = iraisho.write_sheet_name Then
            Set iraisho.write_sheet_obj = Worksheets(i)
            Exit For
        End If
    Next
    If iraisho.write_sheet_obj Is Nothing Then
        Worksheets(iraisho.genpon_sheet_name).Copy _
            after:=Worksheets(iraisho.genpon_sheet_name)
        Set iraisho.write_sheet_obj = ActiveSheet
        iraisho.write_sheet_obj.Name = iraisho.write_sheet_name
    End If
    iraisho.write_sheet_obj.Activate
End Sub
