Imports System.IO
Imports System.Windows.Forms
Imports Rpa00

Module Setup
    Const ICSP As String = "IntranetClientServer"
    Const SAP As String = "StandAlone"
    Const CSP As String = "ClientServer"

    Sub Main()
        Dim sw As StreamWriter
        Dim obj As Object
        Dim pidx As String : Dim pname As String = vbNullString
        Dim yorn As String = vbNullString : Dim yorn2 As String = vbNullString
        Dim flag As Boolean = False
        Dim endflag As Boolean = False

        If Not Directory.Exists(CommonProject.SystemDirectory) Then
            Directory.CreateDirectory(CommonProject.SystemDirectory)
        End If
        If Not Directory.Exists(CommonProject.SystemDllDirectory) Then
            Directory.CreateDirectory(CommonProject.SystemDllDirectory)
        End If

        Do
            Console.WriteLine($"設定項目の選択")
            Console.WriteLine($"    1 ... 新規作成")
            Console.WriteLine($"    2 ... 変更")
            Console.WriteLine($"    3 ... 変更")
        Until endflag

        Do
            Console.WriteLine($"プロジェクト構成を選択")
            Console.WriteLine($"    1 ... {ICSP}")
            Console.WriteLine($"    2 ... {SAP}")
            Console.WriteLine($"    3 ... {CSP}")
            Console.WriteLine($"    9 ... やっぱりやめる")
            Console.Write($">>> ")
            pidx = Console.ReadLine()
        Loop Until pidx = "1" Or pidx = "2" Or pidx = "3" Or pidx = "9"

        If pidx = "1" Then
            obj = New IntranetClientServerProject
            pname = ICSP
        End If
        If pidx = "2" Then
            obj = New StandAloneProject
            pname = SAP
        End If
        If pidx = "3" Then
            obj = New ClientServerProject
            pname = CSP
        End If
        If pidx = "9" Then
            Exit Sub
        End If

        If Not Directory.Exists(obj.SystemProjectDirectory) Then
            Directory.CreateDirectory(obj.SystemProjectDirectory)
            Console.WriteLine($"ディレクトリ '{obj.SystemProjectDirectory}' を新規作成しました")
        End If
        If Not File.Exists(obj.SystemJsonFileName) Then
            obj.Save(obj.SystemJsonFileName, obj)
            Console.WriteLine($"ファイル     '{obj.SystemJsonFileName}' を新規作成しました")
        End If

        obj = obj.Load(obj.SystemJsonFileName)
        Call _CreateMyDirectory(obj)

        Dim ini As New RpaInitializer
        If File.Exists(CommonProject.SystemIniFileName) Then
            ini = ini.Load(CommonProject.SystemIniFileName)
        End If
        ini.CurrentProjectArchitecture = pname
        Call ini.Save(CommonProject.SystemIniFileName, ini)

        Console.WriteLine($"プロジェクト構成を '{ini.CurrentProjectArchitecture}' に設定しました")
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
