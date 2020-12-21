Imports System.IO
Imports System.Windows.Forms
Imports Rpa00

Module Setup
    Const ICSP As String = "IntranetClientServer"
    Const SAP As String = "StandAlone"
    Const CSP As String = "ClientServer"

    Private _ExecutableStrings As Dictionary(Of String, Func(Of Integer))
    Public ReadOnly Property ExecutableStrings As Dictionary(Of String, Func(Of Integer))
        Get
            If _ExecutableStrings Is Nothing Then
                _ExecutableStrings = New Dictionary(Of String, Func(Of Integer))
                _ExecutableStrings.Add("NewProject", AddressOf NewProject)
            End If
            Return _ExecutableStrings
        End Get
    End Property

    Private Function NewProject() As Integer
        Dim pidx As Integer = 0
        Dim obj As Object
        Dim [sln] As String = vbNullString
        Dim flag As Boolean = False
        Dim yorn As String = vbNullString
        Dim arch As String = vbNullString

        Do
            Console.WriteLine($"プロジェクト構成を選択")
            Console.WriteLine($"            1 ... {ICSP}")
            Console.WriteLine($"            2 ... {SAP}")
            Console.WriteLine($"            3 ... {CSP}")
            Console.WriteLine($"            9 ... やっぱりやめる")
            Console.Write($"NewProject>")
            pidx = Console.ReadLine()
            Console.WriteLine(vbNullString)
        Loop Until pidx = "1" Or pidx = "2" Or pidx = "3" Or pidx = "9"

        If pidx = "1" Then
            obj = New IntranetClientServerProject
            arch = ICSP
        End If
        If pidx = "2" Then
            obj = New StandAloneProject
            arch = SAP
        End If
        If pidx = "3" Then
            obj = New ClientServerProject
            arch = CSP
        End If
        If pidx = "9" Then
            Return 1000
        End If


        Dim ini As New RpaInitializer
        If File.Exists(CommonProject.SystemIniFileName) Then
            ini = ini.Load(CommonProject.SystemIniFileName)
        End If

        flag = False
        Do
            Console.WriteLine($"プロジェクト名を入力してください ...")
            Console.Write($"NewProject>")
            [sln] = Console.ReadLine()
            obj.SolutionName = [sln]
            If Directory.Exists(obj.SystemSolutionDirectory) Then
                Console.WriteLine($"プロジェクト '{obj.SolutionName}' は存在します")
                Console.WriteLine($"他のプロジェクト名を入力してください")
                Console.WriteLine(vbNullString)
                Continue Do
            End If
            If ini.Solutions.Contains(obj.SystemSolutionDirectory) Then
                Do
                    yorn = vbNullString
                    Console.WriteLine($"プロジェクト '{obj.SolutionName}' は存在します")
                    Console.WriteLine($"プロジェクトを上書きしますか(y/n)")
                    Console.Write($"NewProject>")
                    yorn = Console.ReadLine()
                    Console.WriteLine(vbNullString)
                Loop Until yorn = "y" Or yorn = "n"
                If yorn = "n" Then
                    Continue Do
                End If
            End If
            Do
                yorn = vbNullString
                Console.WriteLine($"'{obj.SolutionName}' よろしいですか？(y/n)")
                Console.Write($"NewProject>")
                yorn = Console.ReadLine()
                Console.WriteLine(vbNullString)
            Loop Until yorn = "y" Or yorn = "n"

            If yorn = "y" Then
                Directory.CreateDirectory(obj.SystemSolutionDirectory)
                obj.Save(obj.SystemJsonFileName, obj)
                ini.CurrentSolutionArchitecture = arch
                ini.CurrentSolution = obj.SystemSolutionDirectory
                ini.Solutions.Add(obj.SystemSolutionDirectory)
                Call ini.Save(CommonProject.SystemIniFileName, ini)

                Console.WriteLine($"ディレクトリ '{obj.SystemSolutionDirectory}' を新規作成しました")
                Console.WriteLine(vbNullString)
                flag = True
            Else
                Console.WriteLine($"プロジェクトの新規作成を中止しました")
                Console.WriteLine(vbNullString)
                flag = True
            End If
        Loop Until flag

        Console.WriteLine($"プロジェクト構成を '{ini.CurrentSolutionArchitecture}' に設定しました")
        Return 0
    End Function

    Public Sub Main()
        Dim pname As String = vbNullString
        Dim yorn As String = vbNullString : Dim yorn2 As String = vbNullString
        Dim flag As Boolean = False
        Dim endflag As Boolean = False
        Dim txt As String = vbNullString
        Dim i As Integer
        Dim fnc As Func(Of Integer)

        If Not Directory.Exists(CommonProject.SystemDirectory) Then
            Directory.CreateDirectory(CommonProject.SystemDirectory)
        End If
        If Not Directory.Exists(CommonProject.SystemDllDirectory) Then
            Directory.CreateDirectory(CommonProject.SystemDllDirectory)
        End If

        Do
            Console.Write($"Setup>")
            txt = Console.ReadLine()
            fnc = ExecutableStrings(txt)
            i = fnc()
            Console.WriteLine(vbNullString)
        Loop Until endflag

        Console.WriteLine($"プロジェクト設定変更を終了します")
        Console.ReadLine()
    End Sub


    Private Sub _CreateMyDirectory(ByRef obj As Object)
        Dim yorn As String = vbNullString
        Dim yorn2 As String = vbNullString
        Dim fbd As FolderBrowserDialog
        Do
            yorn = vbNullString
            Console.WriteLine($"MyDirectory の設定を行いますか (y/n)")
            Console.Write($">>> ")
            yorn = Console.ReadLine()
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
                    Console.Write($">>> ")
                    yorn2 = Console.ReadLine()
                Else
                    yorn2 = "x"
                End If
            Loop Until yorn2 = "y" Or yorn2 = "x"
            If yorn2 = "y" Then
                obj.MyDirectory = fbd.SelectedPath
                obj.Save(obj.SystemJsonFileName, obj)
                Console.WriteLine($"MyDirectory が設定されました")
            End If
            If yorn2 = "x" Then
                Console.WriteLine($"MyDirectory の設定は行いませんでした")
            End If
        End If
    End Sub
End Module
