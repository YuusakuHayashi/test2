Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms

Public Class NewProjectCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecutableProjectArchitectures As Integer()
        Get
            Return {
                (New IntranetClientServerProject).SystemArchType,
                (New StandAloneProject).SystemArchType,
                (New ClientServerProject).SystemArchType
            }
        End Get
    End Property

    Public Overrides ReadOnly Property CanExecute(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
        Get
            Return {0, 99}
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableUserLevel As Integer
        Get
            Return RpaCodes.ProjectUserLevel.User
        End Get
    End Property

    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        Dim rpa2 As Object
        Dim flag As Boolean = False
        Dim yorn As String = vbNullString
        Dim pname As String

        rpa2 = _SelectProjectArchitecture()
        If rpa2 Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine(vbNullString)
            trn.ExitFlag = True
            Return 1000
        End If

        rpa2 = _GetSolution(rpa2)
        If rpa2 Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine(vbNullString)
            trn.ExitFlag = True
            Return 1000
        End If

        rpa2 = _GetMyDirectory(rpa2)
        If rpa2 Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine(vbNullString)
            trn.ExitFlag = True
            Return 1000
        End If

        ini = _RegisterSolution(ini, rpa2)
        If ini Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine(vbNullString)
            trn.ExitFlag = True
            Return 1000
        End If

        Directory.CreateDirectory(rpa2.SystemSolutionDirectory)
        Console.WriteLine($"ディレクトリ '{rpa2.SystemSolutionDirectory}' を新規作成しました")
        Call rpa2.Save(rpa2.SystemJsonFileName, rpa2)
        Console.WriteLine($"ファイル     '{rpa2.SystemJsonFileName}' を新規作成しました")
        Call ini.Save(CommonProject.SystemIniFileName, ini)
        Console.WriteLine($"ファイル     '{CommonProject.SystemIniFileName}' を新規作成しました")
        Return 0
    End Function

    Private Function _SelectProjectArchitecture() As Object
        Dim inp As String = vbNullString
        Dim pidx As Integer = 0
        Dim yorn As String = vbNullString
        Dim rpa As Object
        Dim arch(4) As String

        Dim asm As Assembly = Nothing
        Dim [mod] As [Module] = Nothing
        'asm = Assembly.LoadFrom(CommonProject.System00DllFileName)
        asm = Assembly.LoadFrom("C:\Users\yuusa\project\test2\Rpa00\obj\Debug\Rpa00.dll")
        [mod] = asm.GetModule("Rpa00.dll")
        Dim ics_type As Type = [mod].GetType("Rpa00.IntranetClientServerProject")
        Dim sap_type As Type = [mod].GetType("Rpa00.StandAloneProject")
        Dim csp_type As Type = [mod].GetType("Rpa00.ClientServerProject")
        Dim ics = Activator.CreateInstance(ics_type)
        Dim sap = Activator.CreateInstance(sap_type)
        Dim csp = Activator.CreateInstance(csp_type)

        Do
            yorn = vbNullString
            Do
                inp = vbNullString
                Console.WriteLine($"プロジェクト構成を選択")
                Console.WriteLine($"           {ics.SystemArchType} ... {ics.SystemArchTypeName}")
                Console.WriteLine($"           {sap.SystemArchType} ... {sap.SystemArchTypeName}")
                Console.WriteLine($"           {csp.SystemArchType} ... {csp.SystemArchTypeName}")
                Console.WriteLine($"           9 ... やっぱりやめる")
                Console.Write($"setup>")
                inp = Console.ReadLine()
                pidx = inp.ToString()
                Console.WriteLine(vbNullString)
            Loop Until pidx < 10

            If pidx = ics.SystemArchType Then
                rpa = ics
            End If
            If pidx = sap.SystemArchType Then
                rpa = sap
            End If
            If pidx = csp.SystemArchType Then
                rpa = csp
            End If
            If pidx = 9 Then
                rpa = Nothing
                Exit Do
            End If

            If Not Directory.Exists(rpa.SystemArchDirectory) Then
                Directory.CreateDirectory(rpa.SystemArchDirectory)
            End If

            Do
                Console.WriteLine($"よろしいですか？ '{inp} ... {rpa.SystemArchTypeName}' (y/n)")
                Console.Write($"setup>")
                yorn = Console.ReadLine()
                Console.WriteLine(vbNullString)
            Loop Until yorn = "y" Or yorn = "n"
        Loop Until yorn = "y"

        Return rpa
    End Function

    ' rpaプロジェクトオブジェクトの[SolutionName]を更新する
    Private Function _GetSolution(ByRef rpa As Object) As Object
        Dim [sln] As String = vbNullString
        Dim flag As Boolean = False
        Dim yorn As String = vbNullString

        Do
            flag = False
            [sln] = vbNullString

            Console.WriteLine($"プロジェクト名を入力してください ...")
            Console.Write($"setup>")
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
                Console.Write($"setup>")
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
End Class
