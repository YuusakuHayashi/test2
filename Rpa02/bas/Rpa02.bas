Attribute VB_Name = "Rpa02"

Option Base 0

'README

'Dim OUTPUT_BOOK_MULTI As String
'Dim OUTPUT_BOOK_SINGLE As String
'Dim input_book         As String

'Private Type input_sheet_data
''   HISTORY
''      2020/02/21 : 新規作成
'    name                           As String
'    obj                            As Worksheet
'    zinkenhi_cost_cell             As Range
'    keihi_cost_cell                As Range
'    bumonhi_cost_cell              As Range
'    bumonrieki_income_cell         As Range
'    honsyahi_cost_cell             As Range
'    keijorieki_income_cell         As Range
'    arari_income_per_man_cell      As Range
'    bumonrieki_income_per_man_cell As Range
'    roudoubunpairitu_per_cell      As Range
'    zinin_amount_cell              As Range
'End Type
'
'Private Type input_book_data
''   HISTORY
''      2020/02/21 : 新規作成
'    name  As String
'    obj   As Workbook
'    sheet As input_sheet_data
'End Type
'
'Private Type output_sheet_data
''   HISTORY
''      2020/02/21 : 新規作成
'    name                           As String
'    obj                            As Worksheet
'    output_range                   As Range
'    inputdate_cell                 As Range
'    nengetu_cell                   As Range
'    zinkenhi_cost_cell             As Range
'    zinkenhi_kbn_cell              As Range
'    keihi_cost_cell                As Range
'    keihi_kbn_cell                 As Range
'    bumonhi_cost_cell              As Range
'    bumonhi_kbn_cell               As Range
'    bumonrieki_income_cell         As Range
'    bumonrieki_kbn_cell            As Range
'    honsyahi_cost_cell             As Range
'    honsyahi_kbn_cell              As Range
'    keijorieki_income_cell         As Range
'    keijorieki_kbn_cell            As Range
'    arari_income_per_man_cell      As Range
'    arari_kbn_per_man_cell         As Range
'    bumonrieki_income_per_man_cell As Range
'    bumonrieki_kbn_per_man_cell    As Range
'    roudoubunpairitu_per_cell      As Range
'    roudoubunpairitu_kbn_cell      As Range
'    zinin_amount_cell              As Range
'    zinin_kbn_cell                 As Range
'    clear_range                    As Range
'End Type
'
'Private Type output_book_data
''   HISTORY
''      2020/02/21 : 新規作成
'    name  As String
'    obj   As Workbook
'    sheet As output_sheet_data
'End Type
'
'
'Private Type business_data
''   HISTORY
''      2020/02/21 : 新規作成
'    month        As Integer
'    input_book   As input_book_data
'    output_book  As output_book_data
'    loop_counter As Integer
'End Type
'
'Sub test()
'    Call main(3)
'End Sub

Private Sub CreatePunchData(ByRef v() As Variant)
    Stop
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    ' インプット設定
    '-------------------------------------------------------------------------'
    Dim inf As String: inf = v(0)
    Dim ibook As Workbook
    Dim isheet As Worksheet
    Set ibook = Workbooks.Open(inf, 0)
    Set isheet = ibook.ActiveSheet
    '-------------------------------------------------------------------------'
    ' アウトプット設定
    '-------------------------------------------------------------------------'
    Dim otf As String: otf = v(1)
    Dim obook As Workbook
    Dim osheet As Worksheet
    Set obook = Workbooks.Open(otf, 0)
    Set osheet = obook.ActiveSheet
    '-------------------------------------------------------------------------'

    ' 終了設定
    '-------------------------------------------------------------------------'
    obook.Save
    Set osheet = Nothing
    obook.Close
    Set obook = Nothing

    Set isheet = Nothing
    ibook.Close
    Set ibook = Nothing
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    '-------------------------------------------------------------------------'
End Sub

'Private Sub main(ParamArray pa_month() As Variant)
''   HISTORY
''      2020/02/21 : 新規作成
'
'    Dim fso As Object
'    Dim project_dir As String
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    project_dir = Replace(ThisWorkbook.Path, "\" + fso.GetFolder(ThisWorkbook.Path).name, "")
'
'    input_book = project_dir + "\" + "DO-74.xlsx"
'    OUTPUT_BOOK_SINGLE = project_dir + "\" + "輸入石油　計画実績目標マスタ用紙（単月）新用紙.xls"
'    OUTPUT_BOOK_MULTI = project_dir + "\" + "輸入石油　計画実績目標マスタ用紙（複数月）新用紙.xls"
'    
'    Dim bd As business_data
'
''-- set book obj ---------------------------------------------------*
'    Set bd.input_book.obj = Workbooks.Open(input_book, 0)
'
'    If UBound(pa_month) >= 1 Then
'       bd.output_book.name = OUTPUT_BOOK_MULTI
'    Else
'       bd.output_book.name = OUTPUT_BOOK_SINGLE
'    End If
'    Set bd.output_book.obj = Workbooks.Open(bd.output_book.name, 0)
''-------------------------------------------------------------------*
'
''-- sheet clear ----------------------------------------------------*
'    'If bd.output_book.name = OUTPUT_BOOK_MULTI Then
'    '    For j = 1 To bd.input_book.obj.Worksheets.Count
'    '        bd.input_book.sheet.name = bd.input_book.obj.Worksheets(j).name
'    '        Set bd.input_book.sheet.obj = bd.input_book.obj.Worksheets(bd.input_book.sheet.name)
'    '        Call set_sheet_obj(bd)
'    '        Call set_range_obj(bd)
'    '        Call data_clear(bd)
'    '    Next
'    'End If
''-------------------------------------------------------------------*
'
''-- data write -----------------------------------------------------*
'    For bd.loop_counter = LBound(pa_month) To UBound(pa_month)
'        bd.month = CInt(pa_month(bd.loop_counter))
'        For j = 1 To bd.input_book.obj.Worksheets.Count
'            If Not bd.input_book.obj.Worksheets(j).name = "DO74-TOTAL" Then
'                bd.input_book.sheet.name = bd.input_book.obj.Worksheets(j).name
'                Call set_sheet_obj(bd)
'                Call set_range_obj(bd)
'                Call data_write(bd)
'            End If
'        Next
'    Next
''-------------------------------------------------------------------*
'    bd.output_book.obj.Save
'    bd.output_book.obj.Close
'    bd.input_book.obj.Application.DisplayAlerts = False
'    bd.input_book.obj.Close
'    Set bd.output_book.obj = Nothing
'    Set bd.input_book.obj = Nothing
'    Set fso = Nothing
'End Sub
'
'Private Sub data_clear(ByRef bd As business_data)
''   HISTORY
''      2020/02/21 : 新規作成
'    Stop
'    bd.output_book.sheet.clear_range.ClearContents
'End Sub
'
'Private Function get_nengetu(ByRef bd As business_data) As String
'    Dim yy As String
'    Dim mm As String
'
'    Select Case CInt(Year(Now())) + bd.month
'        Case Is <= CInt(Year(Now())) + CInt(month(Now()))
'            yy = StrConv(Str(Year(Now()) Mod 100), vbWide)
'            mm = StrConv(Format(bd.month, "00"), vbWide)
'        Case Else
'            yy = StrConv(Str((Year(Now()) - 1) Mod 100), vbWide)
'            mm = StrConv(Format(bd.month, "00"), vbWide)
'    End Select
'
'    get_nengetu = yy + "／" + mm
'End Function
'
'Private Function get_inputdate(ByRef bd As business_data) As String
'    Dim yy As String
'    Dim mm As String
'    Dim dd As String
'    yy = StrConv(Str(Year(Now()) Mod 100), vbWide)
'    mm = StrConv(Str(month(Now())), vbWide)
'    dd = StrConv(Str(Day(Now())), vbWide)
'    get_inputdate = "記入日： " + yy + "年 " + mm + "月 " + dd + "日"
'End Function
'
'Private Function check_nodata(ByVal cd As Variant) As Boolean: check_nodata = True
'    If cd = 0 Then
'        check_nodata = False
'        Exit Function
'    End If
'    If cd = vbNullString Then
'        check_nodata = False
'        Exit Function
'    End If
'End Function
'
'Private Sub data_write(ByRef bd As business_data)
''   HISTORY
''      2020/02/21 : 新規作成
'    With bd.output_book.sheet
'        .inputdate_cell.Value = get_inputdate(bd)
'        .nengetu_cell.Value = get_nengetu(bd)
'
'        bd.output_book.sheet.zinkenhi_kbn_cell.ClearContents
'        bd.output_book.sheet.zinkenhi_cost_cell.ClearContents
'        If check_nodata(bd.input_book.sheet.zinkenhi_cost_cell.Value) Then
'            .zinkenhi_kbn_cell.Value = 2
'            .zinkenhi_cost_cell.Value = bd.input_book.sheet.zinkenhi_cost_cell.Value
'        End If
'
'        bd.output_book.sheet.keihi_kbn_cell.ClearContents
'        bd.output_book.sheet.keihi_cost_cell.ClearContents
'        If check_nodata(bd.input_book.sheet.keihi_cost_cell.Value) Then
'            .keihi_kbn_cell.Value = 2
'            .keihi_cost_cell.Value = bd.input_book.sheet.keihi_cost_cell.Value
'        End If
'
'        bd.output_book.sheet.bumonhi_kbn_cell.ClearContents
'        bd.output_book.sheet.bumonhi_cost_cell.ClearContents
'        If check_nodata(bd.input_book.sheet.bumonhi_cost_cell.Value) Then
'            .bumonhi_kbn_cell.Value = 2
'            .bumonhi_cost_cell.Value = bd.input_book.sheet.bumonhi_cost_cell.Value
'        End If
'
'        bd.output_book.sheet.bumonrieki_kbn_cell.ClearContents
'        bd.output_book.sheet.bumonrieki_income_cell.ClearContents
'        If check_nodata(bd.input_book.sheet.bumonrieki_income_cell.Value) Then
'            .bumonrieki_kbn_cell.Value = 2
'            .bumonrieki_income_cell.Value = bd.input_book.sheet.bumonrieki_income_cell.Value
'        End If
'
'        bd.output_book.sheet.honsyahi_kbn_cell.ClearContents
'        bd.output_book.sheet.honsyahi_cost_cell.ClearContents
'        If check_nodata(bd.input_book.sheet.honsyahi_cost_cell.Value) Then
'            .honsyahi_kbn_cell.Value = 2
'            .honsyahi_cost_cell.Value = bd.input_book.sheet.honsyahi_cost_cell.Value
'        End If
'
'        bd.output_book.sheet.keijorieki_kbn_cell.ClearContents
'        bd.output_book.sheet.keijorieki_income_cell.ClearContents
'        If check_nodata(bd.input_book.sheet.keijorieki_income_cell.Value) Then
'            .keijorieki_kbn_cell.Value = 2
'            .keijorieki_income_cell.Value = bd.input_book.sheet.keijorieki_income_cell.Value
'        End If
'
'        bd.output_book.sheet.arari_kbn_per_man_cell.ClearContents
'        bd.output_book.sheet.arari_income_per_man_cell.ClearContents
'        If check_nodata(bd.input_book.sheet.arari_income_per_man_cell.Value) Then
'            .arari_kbn_per_man_cell.Value = 2
'            .arari_income_per_man_cell.Value = bd.input_book.sheet.arari_income_per_man_cell.Value
'        End If
'
'        bd.output_book.sheet.bumonrieki_kbn_per_man_cell.ClearContents
'        bd.output_book.sheet.bumonrieki_income_per_man_cell.ClearContents
'        If check_nodata(bd.input_book.sheet.bumonrieki_income_per_man_cell.Value) Then
'            .bumonrieki_kbn_per_man_cell.Value = 2
'            .bumonrieki_income_per_man_cell.Value = bd.input_book.sheet.bumonrieki_income_per_man_cell.Value
'        End If
'
'        bd.output_book.sheet.roudoubunpairitu_kbn_cell.ClearContents
'        bd.output_book.sheet.roudoubunpairitu_per_cell.ClearContents
'        If check_nodata(bd.input_book.sheet.roudoubunpairitu_per_cell.Value) Then
'            .roudoubunpairitu_kbn_cell.Value = 2
'            .roudoubunpairitu_per_cell.Value = bd.input_book.sheet.roudoubunpairitu_per_cell.Value
'        End If
'
'        bd.output_book.sheet.zinin_kbn_cell.ClearContents
'        bd.output_book.sheet.zinin_amount_cell.ClearContents
'        If check_nodata(bd.input_book.sheet.zinin_amount_cell.Value) Then
'            .zinin_kbn_cell.Value = 2
'            .zinin_amount_cell.Value = bd.input_book.sheet.zinin_amount_cell.Value
'        End If
'    End With
'End Sub
'
'Private Sub set_sheet_obj(ByRef bd As business_data)
''   HISTORY
''      2020/02/21 : 新規作成
'    Set bd.input_book.sheet.obj = bd.input_book.obj.Worksheets(bd.input_book.sheet.name)
'    Select Case bd.input_book.sheet.name
'        Case "DO74-33190": bd.output_book.sheet.name = "東和町"
'        Case "DO74-33191": bd.output_book.sheet.name = "飯田"
'        Case "DO74-33192": bd.output_book.sheet.name = "飯田橋"
'        Case "DO74-33854": bd.output_book.sheet.name = "松川"
'        Case "DO74-13664": bd.output_book.sheet.name = "羽場町"
'        Case "DO74-87001": bd.output_book.sheet.name = "駒ヶ根"
'        Case "DO74-33859": bd.output_book.sheet.name = "営業課"
'        Case "DO74-33862": bd.output_book.sheet.name = "本社"
'        Case "DO74-33864": bd.output_book.sheet.name = "販売店"
'        Case "DO74-33865": bd.output_book.sheet.name = "他社代行・提携SBC"
'        Case "DO74-33866": bd.output_book.sheet.name = "介護桜町"
'        Case "DO74-TOTAL"
'    End Select
'    Set bd.output_book.sheet.obj = bd.output_book.obj.Worksheets(bd.output_book.sheet.name)
'End Sub
'
'Private Sub set_range_obj(ByRef bd As business_data)
''   HISTORY
''      2020/02/21 : 新規作成
'    Select Case bd.month
'        Case 1: col = 16
'        Case 2: col = 17
'        Case 3: col = 18
'        Case 4: col = 4
'        Case 5: col = 5
'        Case 6: col = 6
'        Case 7: col = 8
'        Case 8: col = 9
'        Case 9: col = 10
'        Case 10: col = 12
'        Case 11: col = 13
'        Case 12: col = 14
''------ ﾃﾞｰﾀｸﾘｱ ｼﾞ ﾊ ﾚﾂｲﾁ ｦ ｻﾝｼｮｳ ｼﾅｲ ﾉﾃﾞ ﾃｷﾄｳ ﾅ ｱﾀｲ ｲﾚﾙ -------------------------------------*
'        Case Else: col = 99
''---------------------------------------------------------------------------------------------*
'    End Select
'    Set bd.input_book.sheet.zinkenhi_cost_cell = bd.input_book.sheet.obj.Cells(53, col)
'    Set bd.input_book.sheet.keihi_cost_cell = bd.input_book.sheet.obj.Cells(54, col)
'    Set bd.input_book.sheet.bumonhi_cost_cell = bd.input_book.sheet.obj.Cells(55, col)
'    Set bd.input_book.sheet.bumonrieki_income_cell = bd.input_book.sheet.obj.Cells(56, col)
'    Set bd.input_book.sheet.honsyahi_cost_cell = bd.input_book.sheet.obj.Cells(58, col)
'    Set bd.input_book.sheet.keijorieki_income_cell = bd.input_book.sheet.obj.Cells(59, col)
'    Set bd.input_book.sheet.arari_income_per_man_cell = bd.input_book.sheet.obj.Cells(61, col)
'    Set bd.input_book.sheet.bumonrieki_income_per_man_cell = bd.input_book.sheet.obj.Cells(62, col)
'    Set bd.input_book.sheet.roudoubunpairitu_per_cell = bd.input_book.sheet.obj.Cells(63, col)
'    Set bd.input_book.sheet.zinin_amount_cell = bd.input_book.sheet.obj.Cells(64, col)
'
'    Select Case bd.output_book.name
'        Case OUTPUT_BOOK_SINGLE
'            With bd.output_book.sheet
'                Set .inputdate_cell = .obj.Cells(4, 10)
'                Set .nengetu_cell = .obj.Cells(5, 5)
'                Set .zinkenhi_kbn_cell = .obj.Cells(59, 5)
'                Set .zinkenhi_cost_cell = .obj.Cells(59, 7)
'                Set .keihi_kbn_cell = .obj.Cells(60, 5)
'                Set .keihi_cost_cell = .obj.Cells(60, 7)
'                Set .bumonhi_kbn_cell = .obj.Cells(61, 5)
'                Set .bumonhi_cost_cell = .obj.Cells(61, 7)
'                Set .bumonrieki_kbn_cell = .obj.Cells(62, 5)
'                Set .bumonrieki_income_cell = .obj.Cells(62, 7)
'                Set .honsyahi_kbn_cell = .obj.Cells(63, 5)
'                Set .honsyahi_cost_cell = .obj.Cells(63, 7)
'                Set .keijorieki_kbn_cell = .obj.Cells(64, 5)
'                Set .keijorieki_income_cell = .obj.Cells(64, 7)
'                Set .arari_kbn_per_man_cell = .obj.Cells(65, 5)
'                Set .arari_income_per_man_cell = .obj.Cells(65, 7)
'                Set .bumonrieki_kbn_per_man_cell = .obj.Cells(66, 5)
'                Set .bumonrieki_income_per_man_cell = .obj.Cells(66, 7)
'                Set .roudoubunpairitu_kbn_cell = .obj.Cells(67, 5)
'                Set .roudoubunpairitu_per_cell = .obj.Cells(67, 7)
'                Set .zinin_kbn_cell = .obj.Cells(69, 5)
'                Set .zinin_amount_cell = .obj.Cells(69, 7)
'            End With
'        Case OUTPUT_BOOK_MULTI
'            With bd.output_book.sheet
'               Set .inputdate_cell = .obj.Cells(4, 11)
'               Set .nengetu_cell = .obj.Cells(6, bd.loop_counter + 6)
'               Set .zinkenhi_kbn_cell = .obj.Cells(59, 5)
'               Set .zinkenhi_cost_cell = .obj.Cells(59, bd.loop_counter + 6)
'               Set .keihi_kbn_cell = .obj.Cells(60, 5)
'               Set .keihi_cost_cell = .obj.Cells(60, bd.loop_counter + 6)
'               Set .bumonhi_kbn_cell = .obj.Cells(61, 5)
'               Set .bumonhi_cost_cell = .obj.Cells(61, bd.loop_counter + 6)
'               Set .bumonrieki_kbn_cell = .obj.Cells(62, 5)
'               Set .bumonrieki_income_cell = .obj.Cells(62, bd.loop_counter + 6)
'               Set .honsyahi_kbn_cell = .obj.Cells(63, 5)
'               Set .honsyahi_cost_cell = .obj.Cells(63, bd.loop_counter + 6)
'               Set .keijorieki_kbn_cell = .obj.Cells(64, 5)
'               Set .keijorieki_income_cell = .obj.Cells(64, bd.loop_counter + 6)
'               Set .arari_kbn_per_man_cell = .obj.Cells(65, 5)
'               Set .arari_income_per_man_cell = .obj.Cells(65, bd.loop_counter + 6)
'               Set .bumonrieki_kbn_per_man_cell = .obj.Cells(66, 5)
'               Set .bumonrieki_income_per_man_cell = .obj.Cells(66, bd.loop_counter + 6)
'               Set .roudoubunpairitu_kbn_cell = .obj.Cells(67, 5)
'               Set .roudoubunpairitu_per_cell = .obj.Cells(67, bd.loop_counter + 6)
'               Set .zinin_kbn_cell = .obj.Cells(69, 5)
'               Set .zinin_amount_cell = .obj.Cells(69, bd.loop_counter + 6)
'               Set .clear_range = _
'                   Union(.obj.Range(.obj.Cells(6, 6), .obj.Cells(6, 10)), .obj.Range(.obj.Cells(8, 5), .obj.Cells(69, 10)))
'            End With
'    End Select
'End Sub
