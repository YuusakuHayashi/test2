﻿Imports System.IO
Imports Rpa00

Public Class Rpa02 : Inherits RpaBase(Of Rpa02)
    Private _OriginDirectory As String
    Private ReadOnly Property OriginDirectory As String
        Get
            If String.IsNullOrEmpty(Me._OriginDirectory) Then
                Me._OriginDirectory = $"{Data.Project.MyRobotDirectory}\origin"
                If Not Directory.Exists(Me._OriginDirectory) Then
                    Directory.CreateDirectory(Me._OriginDirectory)
                End If
            End If
            Return Me._OriginDirectory
        End Get
    End Property

    Private _WorkDirectory As String
    Private ReadOnly Property WorkDirectory As String
        Get
            If String.IsNullOrEmpty(Me._WorkDirectory) Then
                Me._WorkDirectory = $"{Data.Project.MyRobotDirectory}\work"
                If Not Directory.Exists(Me._WorkDirectory) Then
                    Directory.CreateDirectory(Me._WorkDirectory)
                End If
            End If
            Return Me._WorkDirectory
        End Get
    End Property

    Public Class InputBookDatas
        Public BookName As String

        ' ここでＡＤＤしてしまうと、MyRobotとRootRobotが同じ参照のため、２４個分作ってしまう
        Private _MonthCellDatas As List(Of CellDatas)
        Public Property MonthCellDatas As List(Of CellDatas)
            Get
                If Me._MonthCellDatas Is Nothing Then
                    Me._MonthCellDatas = New List(Of CellDatas)
                End If
                Return Me._MonthCellDatas
            End Get
            Set(value As List(Of CellDatas))
                Me._MonthCellDatas = value
            End Set
        End Property

        Public Class CellDatas
            Public ZinkenhiValueCell As String
            Public KeihiValueCell As String
            Public BumonhiValueCell As String
            Public BumonriekiValueCell As String
            Public HonsyahiValueCell As String
            Public KeijoriekiValueCell As String
            Public ArariValueCell As String
            Public AverageBumonriekiValueCell As String
            Public RoudoubunpairituValueCell As String
            Public ZininValueCell As String
        End Class
    End Class

    Public Class OutputBookDatas
        Public BookName As String

        Public WrittenDateCell As String
        Public ZinkenhiKindCell As String
        Public KeihiKindCell As String
        Public BumonhiKindCell As String
        Public BumonriekiKindCell As String
        Public HonsyahiKindCell As String
        Public KeijoriekiKindCell As String
        Public ArariKindCell As String
        Public AverageBumonriekiKindCell As String
        Public RoudoubunpairituKindCell As String
        Public ZininKindCell As String

        Public Class ValueCellDatas
            Public NengetsuCell As String
            Public ZinkenhiValueCell As String
            Public KeihiValueCell As String
            Public BumonhiValueCell As String
            Public BumonriekiValueCell As String
            Public HonsyahiValueCell As String
            Public KeijoriekiValueCell As String
            Public ArariValueCell As String
            Public AverageBumonriekiValueCell As String
            Public RoudoubunpairituValueCell As String
            Public ZininValueCell As String
        End Class

        Private _MonthCellDatas As List(Of ValueCellDatas)
        Public Property MonthCellDatas As List(Of ValueCellDatas)
            Get
                If Me._MonthCellDatas Is Nothing Then
                    Me._MonthCellDatas = New List(Of ValueCellDatas)
                End If
                Return Me._MonthCellDatas
            End Get
            Set(value As List(Of ValueCellDatas))
                Me._MonthCellDatas = value
            End Set
        End Property

        Private _InputOutputSheetCorrespondence As Dictionary(Of String, String)
        Public Property InputOutputSheetCorrespondence As Dictionary(Of String, String)
            Get
                If Me._InputOutputSheetCorrespondence Is Nothing Then
                    Me._InputOutputSheetCorrespondence = New Dictionary(Of String, String)
                End If
                Return Me._InputOutputSheetCorrespondence
            End Get
            Set(value As Dictionary(Of String, String))
                Me._InputOutputSheetCorrespondence = value
            End Set
        End Property
    End Class


    Private _DOBookDatas As InputBookDatas
    Public Property DOBookDatas As InputBookDatas
        Get
            If Me._DOBookDatas Is Nothing Then
                Me._DOBookDatas = New InputBookDatas
            End If
            Return Me._DOBookDatas
        End Get
        Set(value As InputBookDatas)
            Me._DOBookDatas = value
        End Set
    End Property

    Private _SingleMonthBookDatas As OutputBookDatas
    Public Property SingleMonthBookDatas As OutputBookDatas
        Get
            If Me._SingleMonthBookDatas Is Nothing Then
                Me._SingleMonthBookDatas = New OutputBookDatas
            End If
            Return Me._SingleMonthBookDatas
        End Get
        Set(value As OutputBookDatas)
            Me._SingleMonthBookDatas = value
        End Set
    End Property

    Private _MultiMonthBookDatas As OutputBookDatas
    Public Property MultiMonthBookDatas As OutputBookDatas
        Get
            If Me._MultiMonthBookDatas Is Nothing Then
                Me._MultiMonthBookDatas = New OutputBookDatas
            End If
            Return Me._MultiMonthBookDatas
        End Get
        Set(value As OutputBookDatas)
            Me._MultiMonthBookDatas = value
        End Set
    End Property

    Public Overrides Function SetupProjectObject(project As String) As Object
        Throw New NotImplementedException()
    End Function

    Private Function Main(ByRef dat As Object) As Integer
        Dim mutil As New RpaMacroUtility

        Dim months As List(Of String) = dat.System.CommandData.Parameters

        ' シートの対応関係表
        'Dim siosc As Dictionary(Of String, String) = Data.Project.RootRobotObject.SingleMonthBookDatas.InputOutputSheetCorrespondence
        'Dim miosc As Dictionary(Of String, String) = Data.Project.RootRobotObject.MultiMonthBookDatas.InputOutputSheetCorrespondence

        ' 入力ファイルチェック (廃止)
        'Console.WriteLine()
        'If Not Rpa00.RpaModule.FileCheckLoop(Me.DOBookDatas.BookName, dat) Then
        '    Console.WriteLine($"中断しました")
        '    Console.WriteLine()
        '    Return 1000
        'End If

        Dim sorgbook As String = $"{Me.SingleMonthBookDatas.BookName}"
        Dim morgbook As String = $"{Me.MultiMonthBookDatas.BookName}"
        Dim sdstbook As String = $"{Data.Project.MyRobotDirectory}\{Path.GetFileName(Me.SingleMonthBookDatas.BookName)}"
        Dim mdstbook As String = $"{Data.Project.MyRobotDirectory}\{Path.GetFileName(Me.MultiMonthBookDatas.BookName)}"
        If File.Exists(sdstbook) Then
            File.Delete(sdstbook)
        End If
        If File.Exists(mdstbook) Then
            File.Delete(mdstbook)
        End If
        If months.Count = 1 Then
            File.Copy(sorgbook, sdstbook, True)
        Else
            File.Copy(morgbook, mdstbook, True)
        End If

        ' 入力シート
        Dim insheets() As String
        If months.Count = 1 Then
            'insheets = siosc.Keys.ToArray
            insheets = Me.SingleMonthBookDatas.InputOutputSheetCorrespondence.Keys.ToArray
        Else
            'insheets = miosc.Keys.ToArray
            insheets = Me.MultiMonthBookDatas.InputOutputSheetCorrespondence.Keys.ToArray
        End If

        ' 入力データ
        ' Excel側に１次元配列しか渡せないので、Excel側でCSVを分割させる
        Dim inposdatas() As String : ReDim inposdatas(months.Count - 1)
        For Each mon In months
            Dim datastring As String = vbNullString
            datastring &= $"{Me.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).ZinkenhiValueCell}"
            datastring &= $",{Me.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).KeihiValueCell}"
            datastring &= $",{Me.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).BumonhiValueCell}"
            datastring &= $",{Me.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).BumonriekiValueCell}"
            datastring &= $",{Me.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).HonsyahiValueCell}"
            datastring &= $",{Me.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).KeijoriekiValueCell}"
            datastring &= $",{Me.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).ArariValueCell}"
            datastring &= $",{Me.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).AverageBumonriekiValueCell}"
            datastring &= $",{Me.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).RoudoubunpairituValueCell}"
            datastring &= $",{Me.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).ZininValueCell}"
            inposdatas(months.IndexOf(mon)) = datastring
        Next

        Dim [j] As Object
        If months.Count = 1 Then
            [j] = mutil.InvokeMacroFunction("Rpa02.GetInputDatas", {Me.DOBookDatas.BookName, insheets, inposdatas})
        Else
            [j] = mutil.InvokeMacroFunction("Rpa02.GetInputDatas", {Me.DOBookDatas.BookName, insheets, inposdatas})
        End If

        Dim rtndatas(,,) As String = CType([j], String(,,))
        Dim indatas() As String : ReDim indatas(rtndatas.GetLength(0) - 1)
        For [x] As Integer = 0 To (rtndatas.GetLength(0) - 1)
            Dim indatasstring As String = vbNullString
            If months.Count = 1 Then
                indatasstring = $"{Me.SingleMonthBookDatas.InputOutputSheetCorrespondence.ElementAt([x]).Value}"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.WrittenDateCell}{vbTab}{GetWrittenDateString()}"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.ZinkenhiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.KeihiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.BumonhiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.BumonriekiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.HonsyahiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.KeijoriekiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.ArariKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.AverageBumonriekiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.RoudoubunpairituKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.ZininKindCell}{vbTab}2"
            Else
                indatasstring = $"{Me.MultiMonthBookDatas.InputOutputSheetCorrespondence.ElementAt([x]).Value}"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.WrittenDateCell}{vbTab}{GetWrittenDateString()}"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.ZinkenhiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.KeihiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.BumonhiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.BumonriekiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.HonsyahiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.KeijoriekiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.ArariKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.AverageBumonriekiKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.RoudoubunpairituKindCell}{vbTab}2"
                indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.ZininKindCell}{vbTab}2"
            End If
            For [y] As Integer = 0 To (rtndatas.GetLength(1) - 1)
                Dim rtn0 As String = rtndatas([x], [y], 0)
                Dim rtn1 As String = rtndatas([x], [y], 1)
                Dim rtn2 As String = rtndatas([x], [y], 2)
                Dim rtn3 As String = rtndatas([x], [y], 3)
                Dim rtn4 As String = rtndatas([x], [y], 4)
                Dim rtn5 As String = rtndatas([x], [y], 5)
                Dim rtn6 As String = rtndatas([x], [y], 6)
                Dim rtn7 As String = rtndatas([x], [y], 7)
                Dim rtn8 As String = rtndatas([x], [y], 8)
                Dim rtn9 As String = rtndatas([x], [y], 9)
                If months.Count = 1 Then
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).NengetsuCell}{vbTab}{GetNengetsuString(months([y]))}"
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).ZinkenhiValueCell}{vbTab}{rtn0}"
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).KeihiValueCell}{vbTab}{rtn1}"
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).BumonhiValueCell}{vbTab}{rtn2}"
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).BumonriekiValueCell}{vbTab}{rtn3}"
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).HonsyahiValueCell}{vbTab}{rtn4}"
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).KeijoriekiValueCell}{vbTab}{rtn5}"
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).ArariValueCell}{vbTab}{rtn6}"
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).AverageBumonriekiValueCell}{vbTab}{rtn7}"
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).RoudoubunpairituValueCell}{vbTab}{rtn8}"
                    indatasstring &= $"{vbTab}{Me.SingleMonthBookDatas.MonthCellDatas(0).ZininValueCell}{vbTab}{rtn9}"
                Else
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).NengetsuCell}{vbTab}{GetNengetsuString(months([y]))}"
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).ZinkenhiValueCell}{vbTab}{rtn0}"
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).KeihiValueCell}{vbTab}{rtn1}"
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).BumonhiValueCell}{vbTab}{rtn2}"
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).BumonriekiValueCell}{vbTab}{rtn3}"
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).HonsyahiValueCell}{vbTab}{rtn4}"
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).KeijoriekiValueCell}{vbTab}{rtn5}"
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).ArariValueCell}{vbTab}{rtn6}"
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).AverageBumonriekiValueCell}{vbTab}{rtn7}"
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).RoudoubunpairituValueCell}{vbTab}{rtn8}"
                    indatasstring &= $"{vbTab}{Me.MultiMonthBookDatas.MonthCellDatas([y]).ZininValueCell}{vbTab}{rtn9}"
                End If
            Next
            indatas([x]) = indatasstring
        Next

        Dim [k] As Integer = -1
        If months.Count = 1 Then
            [k] = mutil.InvokeMacro("Rpa02.WriteOutputData", {sdstbook, indatas})
        Else
            [k] = mutil.InvokeMacro("Rpa02.WriteOutputData", {mdstbook, indatas})
        End If

        dat.System.RpaWriteLine($"処理完了！")
        dat.System.RpaWriteLine()

        Return 0
    End Function

    Private Function GetWrittenDateString() As String
        Dim dtstring As String = $"記入日： {Date.Now.ToString("yy年 MM月 dd日")}"
        Return dtstring
    End Function

    Private Function GetNengetsuString(ByVal inpmonth As String) As String
        Dim fmtmonth As String = IIf(inpmonth.Length = 2, inpmonth, inpmonth.PadLeft(2, "0"))
        Dim curyear As Integer = Integer.Parse(Date.Now.ToString("yy"))
        Dim inpdate As Integer = curyear * 100 + Integer.Parse(inpmonth)
        Dim curdate As Integer = Integer.Parse(Date.Now.ToString("yyMM"))
        Dim dtstring As String = vbNullString
        If inpdate > curdate Then
            dtstring = $"{(curyear - 1).ToString}／{fmtmonth}"
        Else
            dtstring = $"{(curyear).ToString}／{fmtmonth}"
        End If
        Return dtstring
    End Function

    Private Function Check(ByRef dat As Object) As Boolean
        Dim mutil As New RpaMacroUtility
        Dim outil As New RpaOutlookUtility

        Dim [new] As List(Of String)
        If dat.System.CommandData.Parameters.Contains("nocheck") Then
            'dat.System.CommandData.Parameters = Rpa00.RpaModule.Pop(Of List(Of String))(dat.System.CommandData.Parameters)
            [new] = dat.System.CommandData.Parameters
            [new].Remove("nocheck")
            dat.System.CommandData.Parameters = [new]
            Return True
        End If

        If String.IsNullOrEmpty(Me.DOBookDatas.BookName) Then
            dat.System.RpaWriteLine($" 'DOBookDatas.BookName' が設定されていません")
            Return False
        End If
        If Not File.Exists(Me.DOBookDatas.BookName) Then
            dat.System.RpaWriteLine($"ファイル '{Me.DOBookDatas.BookName}' が存在しません")
            Return False
        End If
        If String.IsNullOrEmpty(Me.SingleMonthBookDatas.BookName) Then
            dat.System.RpaWriteLine($"'SingleMonthBookDatas.BookName' が設定されていません")
            Return False
        End If
        If Not File.Exists(Me.SingleMonthBookDatas.BookName) Then
            dat.System.RpaWriteLine($"ファイル '{Me.SingleMonthBookDatas.BookName}' が存在しません")
            Return False
        End If
        If String.IsNullOrEmpty(Me.MultiMonthBookDatas.BookName) Then
            dat.System.RpaWriteLine($"'MultiMonthBookDatas.BookName' が設定されていません")
            Return False
        End If
        If Not File.Exists(Me.MultiMonthBookDatas.BookName) Then
            dat.System.RpaWriteLine($"ファイル '{Me.MultiMonthBookDatas.BookName}' が存在しません")
            Return False
        End If

        If Not File.Exists(mutil.MacroFileName) Then
            dat.System.RpaWriteLine($"マクロファイル '{mutil.MacroFileName}' が存在しません")
            Return False
        End If

        Dim obj As Object
        Dim ck As Boolean = True
        For Each [mod] In {"Rpa02", "RpaSystem"}
            obj = mutil.InvokeMacroFunction("MacroImporter.IsModuleExist", {[mod]})
            If obj IsNot Nothing Then
                Dim ck2 = CType(obj, Boolean)
                If Not ck2 Then
                    dat.System.RpaWriteLine($"ファイル '{mutil.MacroFileName}' 内にモジュール '{[mod]}' が存在しません")
                    ck = False
                End If
            Else
                dat.System.RpaWriteLine($"モジュールのチェックに失敗しました")
                ck = False
                Exit For
            End If
        Next
        If Not ck Then
            Return False
        End If

        obj = Nothing
        ck = True
        For Each ws In Me.SingleMonthBookDatas.InputOutputSheetCorrespondence.Keys
            obj = mutil.InvokeMacroFunction("RpaSystem.IsSheetExist", {Me.DOBookDatas.BookName, ws})
            If obj IsNot Nothing Then
                Dim ck2 = CType(obj, Boolean)
                If Not ck2 Then
                    dat.System.RpaWriteLine($"'SingleMonthBookDatas.InputOutputSheetCorrespondence' の")
                    dat.System.RpaWriteLine($"シート名 '{ws}' は、 '{Me.DOBookDatas.BookName}' 内に存在しません")
                    ck = False
                End If
            Else
                dat.System.RpaWriteLine($"シートのチェックに失敗しました")
                ck = False
                Exit For
            End If
        Next
        If Not ck Then
            Return False
        End If

        obj = Nothing
        ck = True
        For Each ws In Me.SingleMonthBookDatas.InputOutputSheetCorrespondence.Values
            obj = mutil.InvokeMacroFunction("RpaSystem.IsSheetExist", {Me.SingleMonthBookDatas.BookName, ws})
            If obj IsNot Nothing Then
                Dim ck2 = CType(obj, Boolean)
                If Not ck2 Then
                    dat.System.RpaWriteLine($"'SingleMonthBookDatas.InputOutputSheetCorrespondence' の")
                    dat.System.RpaWriteLine($"シート名 '{ws}' は、 '{Me.SingleMonthBookDatas.BookName}' 内に存在しません")
                    ck = False
                End If
            Else
                dat.System.RpaWriteLine($"シートのチェックに失敗しました")
                ck = False
                Exit For
            End If
        Next
        If Not ck Then
            Return False
        End If

        obj = Nothing
        ck = True
        For Each ws In Me.MultiMonthBookDatas.InputOutputSheetCorrespondence.Keys
            obj = mutil.InvokeMacroFunction("RpaSystem.IsSheetExist", {Me.DOBookDatas.BookName, ws})
            If obj IsNot Nothing Then
                Dim ck2 = CType(obj, Boolean)
                If Not ck2 Then
                    dat.System.RpaWriteLine($"'MultiMonthBookDatas.InputOutputSheetCorrespondence' の")
                    dat.System.RpaWriteLine($"シート名 '{ws}' は、 '{Me.DOBookDatas.BookName}' 内に存在しません")
                    ck = False
                End If
            Else
                dat.System.RpaWriteLine($"シートのチェックに失敗しました")
                ck = False
                Exit For
            End If
        Next
        If Not ck Then
            Return False
        End If

        obj = Nothing
        ck = True
        For Each ws In Me.MultiMonthBookDatas.InputOutputSheetCorrespondence.Values
            obj = mutil.InvokeMacroFunction("RpaSystem.IsSheetExist", {Me.MultiMonthBookDatas.BookName, ws})
            If obj IsNot Nothing Then
                Dim ck2 = CType(obj, Boolean)
                If Not ck2 Then
                    dat.System.RpaWriteLine($"'MultiMonthBookDatas.InputOutputSheetCorrespondence' の")
                    dat.System.RpaWriteLine($"シート名 '{ws}' は、 '{Me.MultiMonthBookDatas.BookName}' 内に存在しません")
                    ck = False
                End If
            Else
                dat.System.RpaWriteLine($"シートのチェックに失敗しました")
                ck = False
                Exit For
            End If
        Next
        If Not ck Then
            Return False
        End If

        If dat.System.CommandData.Parameters.Count = 0 Or dat.System.CommandData.Parameters.Count > 5 Then
            dat.System.RpaWriteLine($"パラメータ数が不正です")
            Return False
        End If

        ck = True
        For Each para In dat.System.CommandData.Parameters
            If Not (para = "1" Or para = "2" Or para = "3" Or para = "4" Or para = "5" Or para = "6" Or para = "7" Or para = "8" Or para = "9" Or para = "10" Or para = "11" Or para = "12") Then
                dat.System.RpaWriteLine($"パラメータ '{para}' は不正です")
                ck = False
            End If
        Next
        If Not ck Then
            Return False
        End If

        Return True
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
    End Sub
End Class
