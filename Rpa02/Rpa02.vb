Imports System.IO
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
        Public NengetsuCell As String
        Public ZinkenhiValueCell As String
        Public ZinkenhiKindCell As String
        Public KeihiValueCell As String
        Public KeihiKindCell As String
        Public BumonhiValueCell As String
        Public BumonhiKindCell As String
        Public BumonriekiValueCell As String
        Public BumonriekiKindCell As String
        Public HonsyahiValueCell As String
        Public HonsyahiKindCell As String
        Public KeijoriekiValueCell As String
        Public KeijoriekiKindCell As String
        Public ArariValueCell As String
        Public ArariKindCell As String
        Public AverageBumonriekiValueCell As String
        Public AverageBumonriekiKindCell As String
        Public RoudoubunpairituValueCell As String
        Public RoudoubunpairituKindCell As String
        Public ZininValueCell As String
        Public ZininKindCell As String

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

    Public Overrides Function Execute(ByRef dat As Object) As Integer
        Dim months As List(Of String) = dat.Transaction.Parameters

        ' 入力ファイルチェック
        If Not Rpa00.RpaModule.FileCheckLoop(Me.DOBookDatas.BookName, dat) Then
            Console.WriteLine($"中断しました")
            Console.WriteLine()
            Return 1000
        End If

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

        ' 入力データの作成
        ' Excel側に１次元配列しか渡せないので、Excel側でCSVを分割する
        Dim indatas() As String : ReDim indatas(months.Count - 1)

        For Each mon In months
            Dim datastring As String = vbNullString
            datastring &= $"{Data.Project.RootRobotObject.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).ZinkenhiValueCell}"
            datastring &= $",{Data.Project.RootRobotObject.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).KeihiValueCell}"
            datastring &= $",{Data.Project.RootRobotObject.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).BumonhiValueCell}"
            datastring &= $",{Data.Project.RootRobotObject.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).BumonriekiValueCell}"
            datastring &= $",{Data.Project.RootRobotObject.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).HonsyahiValueCell}"
            datastring &= $",{Data.Project.RootRobotObject.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).KeijoriekiValueCell}"
            datastring &= $",{Data.Project.RootRobotObject.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).ArariValueCell}"
            datastring &= $",{Data.Project.RootRobotObject.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).AverageBumonriekiValueCell}"
            datastring &= $",{Data.Project.RootRobotObject.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).RoudoubunpairituValueCell}"
            datastring &= $",{Data.Project.RootRobotObject.DOBookDatas.MonthCellDatas(Integer.Parse(mon) - 1).ZininValueCell}"
            indatas(months.IndexOf(mon)) = datastring
        Next

        ' 出力データの作成
        Dim otdatas As String = vbNullString
        If months.Count = 1 Then
            otdatas &= $"{Data.Project.RootRobotObject.SingleMonthBookDatas.WrittenDateCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.NengetsuCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.ZinkenhiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.ZinkenhiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.KeihiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.KeihiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.BumonhiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.BumonhiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.BumonriekiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.BumonriekiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.HonsyahiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.HonsyahiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.KeijoriekiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.KeijoriekiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.ArariValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.ArariKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.AverageBumonriekiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.AverageBumonriekiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.RoudoubunpairituValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.RoudoubunpairituKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.ZininValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.SingleMonthBookDatas.ZininKindCell}"
        Else
            otdatas &= $"{Data.Project.RootRobotObject.MultiMonthBookDatas.WrittenDateCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.NengetsuCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.ZinkenhiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.ZinkenhiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.KeihiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.KeihiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.BumonhiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.BumonhiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.BumonriekiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.BumonriekiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.HonsyahiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.HonsyahiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.KeijoriekiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.KeijoriekiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.ArariValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.ArariKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.AverageBumonriekiValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.AverageBumonriekiKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.RoudoubunpairituValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.RoudoubunpairituKindCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.ZininValueCell}"
            otdatas &= $",{Data.Project.RootRobotObject.MultiMonthBookDatas.ZininKindCell}"
        End If


        Dim mutil As New RpaMacroUtility
        Dim i As Integer = -1
        If months.Count = 1 Then
            i = mutil.InvokeMacro("Rpa02.CreatePunchData", {Me.DOBookDatas.BookName, sdstbook, indatas, otdatas})
        Else
            i = mutil.InvokeMacro("Rpa02.CreatePunchData", {Me.DOBookDatas.BookName, mdstbook, indatas, otdatas})
        End If

        Return 0
    End Function

    Public Overrides Function CanExecute(ByRef dat As Object) As Boolean
        Dim mutil As New RpaMacroUtility
        Dim outil As New RpaOutlookUtility

        If String.IsNullOrEmpty(Me.DOBookDatas.BookName) Then
            Console.WriteLine($"'DOBookDatas.BookName' が設定されていません")
            Return False
        End If
        If String.IsNullOrEmpty(Me.SingleMonthBookDatas.BookName) Then
            Console.WriteLine($"'SingleMonthBookDatas.BookName' が設定されていません")
            Return False
        End If
        If Not File.Exists(Me.SingleMonthBookDatas.BookName) Then
            Console.WriteLine($"ファイル '{Me.SingleMonthBookDatas.BookName}' が存在しません")
            Return False
        End If
        If String.IsNullOrEmpty(Me.MultiMonthBookDatas.BookName) Then
            Console.WriteLine($"'MultiMonthBookDatas.BookName' が設定されていません")
            Return False
        End If
        If Not File.Exists(Me.MultiMonthBookDatas.BookName) Then
            Console.WriteLine($"ファイル '{Me.MultiMonthBookDatas.BookName}' が存在しません")
            Return False
        End If

        If Not File.Exists(mutil.MacroFileName) Then
            Console.WriteLine($"マクロファイル '{mutil.MacroFileName}' が存在しません")
            Return False
        End If

        Dim obj As Object
        Dim ck As Boolean = True
        For Each [mod] In {"Rpa02"}
            obj = mutil.InvokeMacroFunction("MacroImporter.IsModuleExist", {[mod]})
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

        If dat.Transaction.Parameters.Count = 0 Then
            Console.WriteLine($"パラメータ数が不正です")
            Return False
        End If

        Dim ck3 As Boolean = True
        For Each para In dat.Transaction.Parameters
            If Not (para = "1" Or para = "2" Or para = "3" Or para = "4" Or para = "5" Or para = "6" Or para = "7" Or para = "8" Or para = "9" Or para = "10" Or para = "11" Or para = "12") Then
                Console.WriteLine($"パラメータ '{para}' は不正です")
                ck3 = False
            End If
        Next
        If Not ck3 Then
            Return False
        End If

        Return True
    End Function
End Class
