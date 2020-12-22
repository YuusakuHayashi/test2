Imports System.IO
Imports System.Windows.Forms
Imports Rpa00

Module Setup

    Private _ExecutableStrings As Dictionary(Of String, Func(Of RpaInitializer, RpaInitializer))
    Public ReadOnly Property ExecutableStrings As Dictionary(Of String, Func(Of RpaInitializer, RpaInitializer))
        Get
            If _ExecutableStrings Is Nothing Then
                _ExecutableStrings = New Dictionary(Of String, Func(Of RpaInitializer, RpaInitializer))
                _ExecutableStrings.Add("NewProject", AddressOf NewProject)
                _ExecutableStrings.Add("AutoLoad", AddressOf SetAutoLoad)
                _ExecutableStrings.Add("ChangeProject", AddressOf ChangeSolution)
                _ExecutableStrings.Add("Exit", AddressOf ExitSetup)
            End If
            Return _ExecutableStrings
        End Get
    End Property

    ' Select Solution
    '---------------------------------------------------------------------------------------------'
    Private Function ChangeSolution(ini As RpaInitializer) As RpaInitializer
        Dim idx As Integer = 0
        Dim idxstr As String = vbNullString
        Dim idx2 As Integer = -1
        Dim flag As Boolean = False

        For Each sl In ini.Solutions
            Console.WriteLine($"{idx}... '{sl.Name}'")
            idx += 1
        Next
        Console.WriteLine($"現在のプロジェクトは '{ini.CurrentSolution.Name}' です")
        Console.WriteLine(vbNullString)
        Do
            flag = False
            idx2 = -1
            idxstr = vbNullString
            Console.WriteLine($"プロジェクトＮｏを選択")
            Console.Write($"ChangeProject>")
            idxstr = Console.ReadLine()
            Console.WriteLine(vbNullString)

            idx2 = idxstr.ToString()
            If (idx2 + 1) > ini.Solutions.Count Then
                Console.WriteLine($"プロジェクトＮｏが範囲外です")
                Console.WriteLine(vbNullString)
                Continue Do
            Else
                ini.CurrentSolution = ini.Solutions(idx2)
                Console.WriteLine($"プロジェクト '{ini.Solutions(idx2).Name}' に変更しました")
                Console.WriteLine(vbNullString)
                flag = True
            End If
        Loop Until flag

        Return ini
    End Function

    ' New Project
    '---------------------------------------------------------------------------------------------'
    Private Function NewProject(ini As RpaInitializer) As RpaInitializer
        Dim pidx As Integer = 0
        Dim arch As Object
        Dim [sln] As RpaInitializer.RpaSolution
        Dim flag As Boolean = False
        Dim yorn As String = vbNullString
        Dim pname As String

        arch = _SelectProjectArchitecture()
        If arch Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine(vbNullString)
            Return Nothing
        End If

        arch = _GetSolution(arch)
        If arch Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine(vbNullString)
            Return Nothing
        End If

        arch = _GetMyDirectory(arch)
        If arch Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine(vbNullString)
            Return Nothing
        End If

        ini = _RegisterSolution(ini, arch)
        If ini Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine(vbNullString)
            Return Nothing
        End If

        Directory.CreateDirectory(arch.SystemSolutionDirectory)
        Console.WriteLine($"ディレクトリ '{arch.SystemSolutionDirectory}' を新規作成しました")
        Call arch.Save(arch.SystemJsonFileName, arch)
        Console.WriteLine($"ファイル     '{arch.SystemJsonFileName}' を新規作成しました")
        Return ini
    End Function

    Private Function _SelectProjectArchitecture() As Object
        Dim inp As String = vbNullString
        Dim pidx As Integer = 0
        Dim yorn As String = vbNullString
        Dim obj As Object
        Dim arch(4) As String
        Dim ics As New IntranetClientServerProject
        Dim sap As New StandAloneProject
        Dim csp As New ClientServerProject

        Do
            yorn = vbNullString
            Do
                inp = vbNullString
                Console.WriteLine($"プロジェクト構成を選択")
                Console.WriteLine($"           {ics.SystemArchType} ... {ics.SystemArchTypeName}")
                Console.WriteLine($"           {sap.SystemArchType} ... {sap.SystemArchTypeName}")
                Console.WriteLine($"           {csp.SystemArchType} ... {csp.SystemArchTypeName}")
                Console.WriteLine($"           9 ... やっぱりやめる")
                Console.Write($"NewProject>")
                inp = Console.ReadLine()
                pidx = inp.ToString()
                Console.WriteLine(vbNullString)
            Loop Until pidx < 10
            If pidx = ics.SystemArchType Then
                obj = ics
            End If
            If pidx = sap.SystemArchType Then
                obj = sap
            End If
            If pidx = csp.SystemArchType Then
                obj = csp
            End If
            If pidx = 9 Then
                obj = Nothing
                Exit Do
            End If

            If Not Directory.Exists(obj.SystemArchDirectory) Then
                Directory.CreateDirectory(obj.SystemArchDirectory)
            End If

            Do
                Console.WriteLine($"よろしいですか？ '{inp} ... {obj.SystemArchTypeName}' (y/n)")
                Console.Write($"NewProject>")
                yorn = Console.ReadLine()
                Console.WriteLine(vbNullString)
            Loop Until yorn = "y" Or yorn = "n"
        Loop Until yorn = "y"

        Return obj
    End Function

    Private Function _GetSolution(ByRef rpa As Object) As Object
        Dim [sln] As String = vbNullString
        Dim flag As Boolean = False
        Dim yorn As String = vbNullString

        Do
            flag = False
            [sln] = vbNullString

            Console.WriteLine($"プロジェクト名を入力してください ...")
            Console.Write($"NewProject>")
            [sln] = Console.ReadLine()
            rpa.SolutionName = [sln]
            Console.WriteLine(vbNullString)

            If Directory.Exists(rpa.SystemSolutionDirectory) Then
                Console.WriteLine($"プロジェクト '{rpa.SolutionName}' は存在します")
                Console.WriteLine($"他のプロジェクト名を入力してください")
                Console.WriteLine(vbNullString)
                Continue Do
            End If

            Do
                yorn = vbNullString
                Console.WriteLine($"'{rpa.SolutionName}' よろしいですか？(y/n)")
                Console.Write($"NewProject>")
                yorn = Console.ReadLine()
                Console.WriteLine(vbNullString)
            Loop Until yorn = "y" Or yorn = "n"

            If yorn = "y" Then
                flag = True
            Else
                rpa = Nothing
                flag = True
            End If
        Loop Until flag
        Return rpa
    End Function

    Private Function _RegisterSolution(ini As RpaInitializer, rpa As Object) As RpaInitializer
        Dim sl As RpaInitializer.RpaSolution
        Dim [new] As RpaInitializer.RpaSolution
        Dim idx As Integer = -1
        sl = ini.Solutions.Find(
            Function(s)
                Return (s.Name = rpa.SolutionName)
            End Function
        )
        [new] = New RpaInitializer.RpaSolution With {
            .Name = rpa.SolutionName,
            .Architecture = rpa.SystemArchType,
            .SolutionDirectory = rpa.SystemSolutionDirectory
        }
        If sl Is Nothing Then
            ini.Solutions.Add([new])
        Else
            idx = ini.Solutions.IndexOf(sl)
            ini.Solutions(idx) = [new]
        End If
        ini.CurrentSolution = [new]
        Return ini
    End Function

    ' AutoLoad
    '---------------------------------------------------------------------------------------------'
    Private Function SetAutoLoad(ini As RpaInitializer) As RpaInitializer
        Dim yorn As String = vbNullString
        Dim onoff As String = IIf(ini.AutoLoad, "ON", "OFF")
        Dim [switch] As String = IIf(ini.AutoLoad, "OFF", "ON")

        Console.WriteLine($"現在 AutoLoad = {onoff} です")
        Console.WriteLine($"     AutoLoad = {[switch]} に変更しますか？ (y/n)")

        Do
            yorn = vbNullString
            Console.Write($"AutoLoad>")
            yorn = Console.ReadLine()
            Console.WriteLine(vbNullString)
            If yorn = "y" Then
                ini.AutoLoad = (Not ini.AutoLoad)
            End If
        Loop Until yorn = "y" Or yorn = "n"

        Return ini
    End Function

    ' Exit
    '---------------------------------------------------------------------------------------------'
    Private Function ExitSetup(ini As RpaInitializer) As RpaInitializer
        Return Nothing
    End Function

    Public Sub Main()
        Dim pname As String = vbNullString
        Dim yorn As String = vbNullString : Dim yorn2 As String = vbNullString
        Dim flag As Boolean = False
        Dim endflag As Boolean = False
        Dim txt As String = vbNullString
        Dim i As Integer
        Dim fnc As Func(Of RpaInitializer, RpaInitializer)
        Dim ini As RpaInitializer

        If Not Directory.Exists(CommonProject.SystemDirectory) Then
            Directory.CreateDirectory(CommonProject.SystemDirectory)
        End If
        If Not Directory.Exists(CommonProject.SystemDllDirectory) Then
            Directory.CreateDirectory(CommonProject.SystemDllDirectory)
        End If

        ini = New RpaInitializer
        If Not File.Exists(CommonProject.SystemIniFileName) Then
            ini.Save(CommonProject.SystemIniFileName, ini)
        End If
        ini = ini.Load(CommonProject.SystemIniFileName)

        Do
            Console.Write($"Setup>")
            txt = Console.ReadLine()
            If ExecutableStrings.ContainsKey(txt) Then
                fnc = ExecutableStrings(txt)
                ini = fnc(ini)
                If ini Is Nothing Then
                    endflag = True
                Else
                    Call ini.Save(CommonProject.SystemIniFileName, ini)
                    Console.WriteLine($"ファイル     '{CommonProject.SystemIniFileName}' を変更しました")
                End If
            Else
                Console.WriteLine($"'{txt}' は不正です")
            End If
            Console.WriteLine(vbNullString)
        Loop Until endflag

        Console.WriteLine($"プロジェクト設定変更を終了します")
        Console.ReadLine()
    End Sub


    Private Function _GetMyDirectory(ByRef obj As Object) As Object
        Dim yorn As String = vbNullString
        Dim yorn2 As String = vbNullString
        Dim fbd As FolderBrowserDialog
        Do
            yorn = vbNullString
            Console.WriteLine($"MyDirectory の設定を行いますか (y/n)")
            Console.Write($"GetMyDirectory>")
            yorn = Console.ReadLine()
            Console.WriteLine(vbNullString)
        Loop Until yorn = "y" Or yorn = "n"
        If yorn = "y" Then
            Do
                yorn2 = vbNullString
                fbd = New FolderBrowserDialog With {
                    .Description = $"Select MyDirectory",
                    .RootFolder = Environment.SpecialFolder.Desktop,
                    .SelectedPath = Environment.SpecialFolder.Desktop,
                    .ShowNewFolderButton = True
                }
                If fbd.ShowDialog() = DialogResult.OK Then
                    Console.WriteLine($"よろしいですか？ '{fbd.SelectedPath}' (y/n)")
                    Console.Write($"GetMyDirectory>")
                    yorn2 = Console.ReadLine()
                    Console.WriteLine(vbNullString)
                Else
                    yorn2 = vbNullString
                End If
            Loop Until yorn2 = "y" Or yorn2 = vbNullString
            If yorn2 = "y" Then
                obj.MyDirectory = fbd.SelectedPath
                Console.WriteLine($"MyDirectory が設定されました")
                Console.WriteLine(vbNullString)
            Else
                Console.WriteLine($"MyDirectory の設定は行いませんでした")
                Console.WriteLine(vbNullString)
            End If
        End If
        Return obj
    End Function
End Module
