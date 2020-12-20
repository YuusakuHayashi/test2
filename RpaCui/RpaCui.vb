Imports System.Reflection
Imports System.IO
Imports Rpa00

Module RpaCui

    Sub Main()
        Dim asm As Assembly
        Dim [mod] As [Module]
        Dim rpa_type As Type : Dim trn_type As Type : Dim sys_type As Type
        Dim rpa, trn, sys
        Dim ptype As String = vbNullString
        Dim txt As String = vbNullString
        Dim inv As Boolean = False

        If Not File.Exists(CommonProject.SystemIniFileName) Then
            Console.WriteLine($"ファイル     '{CommonProject.SystemIniFileName}' が存在しません")
            Console.ReadLine()
            Exit Sub
        End If
        If Not Directory.Exists(CommonProject.SystemDllDirectory) Then
            Console.WriteLine($"ディレクトリ '{CommonProject.SystemDllDirectory}' が存在しません")
            Console.ReadLine()
            Exit Sub
        End If
        If Not File.Exists(CommonProject.System00DllFileName) Then
            Console.WriteLine($"ファイル     '{CommonProject.System00DllFileName}' が存在しません")
            Console.ReadLine()
            Exit Sub
        End If

        ' IniFile 読み込み
        Dim flag As Boolean = False
        Dim jh As JsonHandler(Of RpaInitializer)
        Dim ini As RpaInitializer
        jh = New JsonHandler(Of RpaInitializer)
        ini = jh.Load(CommonProject.SystemIniFileName)
        If ini Is Nothing Then
            Console.WriteLine($"ファイル     '{CommonProject.SystemIniFileName}' の読み込みに失敗しました")
            Console.ReadLine()
            Exit Sub
        End If
        flag = IIf(ini.CurrentProjectArchitecture = "IntranetClientServer", True, flag)
        flag = IIf(ini.CurrentProjectArchitecture = "StandAlone", True, flag)
        flag = IIf(ini.CurrentProjectArchitecture = "ClientServer", True, flag)
        If Not flag Then
            Console.WriteLine($"プロジェクト構成 '{ini.CurrentProjectArchitecture}' は想定外です")
            Console.ReadLine()
            Exit Sub
        End If

        ' オブジェクト生成
        asm = Assembly.LoadFrom(CommonProject.System00DllFileName)
        [mod] = asm.GetModule("Rpa00.dll")
        trn_type = [mod].GetType("Rpa00.RpaTransaction")
        sys_type = [mod].GetType("Rpa00.RpaSystem")
        Select Case ini.CurrentProjectArchitecture
            Case "IntranetClientServer"
                rpa_type = [mod].GetType("Rpa00.IntranetClientServerProject")
            Case "StandAlone"
                rpa_type = [mod].GetType("Rpa00.StandAloneProject")
            Case "ClientServer"
                rpa_type = [mod].GetType("Rpa00.ClientServerProject")
        End Select
        If rpa_type IsNot Nothing Then
            rpa = Activator.CreateInstance(rpa_type)
        End If
        If trn_type IsNot Nothing Then
            trn = Activator.CreateInstance(trn_type)
        End If
        If sys_type IsNot Nothing Then
            sys = Activator.CreateInstance(sys_type)
        End If

        Call rpa.HelloProject()
        Call rpa.CheckProject()

        ' 実行
        Do Until trn.ExitFlag
            Console.Write($"{rpa.ProjectAlias} > ")
            trn.CommandText = Console.ReadLine()
            Call trn.CreateCommand()
            Call sys.Main(trn, rpa)
            Console.WriteLine(vbNullString)
        Loop
    End Sub

End Module
