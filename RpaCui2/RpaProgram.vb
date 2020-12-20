Imports System
Imports System.IO
Imports System.Reflection
Imports System.IO.StreamReader

Module Program
    Sub Main(args As String())
        Dim asm As Assembly
        Dim [mod] As [Module]
        Dim rpa_type As Type : Dim trn_type As Type : Dim sys_type As Type
        Dim rpa, trn, sys, rpa2
        Dim rpadir As String : Dim dlldir As String : Dim inifil As String
        Dim rpa00dll As String
        Dim ptype As String = vbNullString
        Dim sw As StreamReader
        Dim txt As String = vbNullString
        Dim inv As Boolean = False

        rpadir = Environment.GetEnvironmentVariable("USERPROFILE") & "\rpa_project"
        dlldir = $"{rpadir}\dll"
        If args.Count > 1 Then
            If args(0) = "Debug" Then
                dlldir = args(1)
            End If
        End If
        rpa00dll = $"{dlldir}\Rpa00.dll"
        inifil = $"{rpadir}\rpa.ini"

        ' チェック
        If Not Directory.Exists(rpadir) Then
            Console.WriteLine($"ディレクトリ '{rpadir}' がありません")
            inv = True
        End If
        If Not Directory.Exists(dlldir) Then
            Console.WriteLine($"ディレクトリ '{dlldir}' がありません")
            inv = True
        End If
        If Not File.Exists(inifil) Then
            Console.WriteLine($"ファイル     '{inifil}' がありません")
            inv = True
        End If
        If Not File.Exists(rpa00dll) Then
            Console.WriteLine($"ＤＬＬ       '{rpa00dll}' がありません")
            inv = True
        End If
        If inv Then
            Console.ReadLine()
            Exit Sub
        End If


        ' プロジェクトタイプの決定
        sw = New StreamReader(inifil, Text.Encoding.GetEncoding("Shift-JIS"))
        txt = Strings.Trim(sw.ReadToEnd())
        sw.Close()
        sw.Dispose()

        asm = Assembly.LoadFrom(rpa00dll)
        [mod] = asm.GetModule("Rpa00.dll")
        trn_type = [mod].GetType("Rpa00.RpaTransaction")
        sys_type = [mod].GetType("Rpa00.RpaSystem")
        Select Case txt
            Case "IntranetClientServer"
                rpa_type = [mod].GetType("Rpa00.IntranetClientServerProject")
            Case "StandAlone"
                rpa_type = [mod].GetType("Rpa00.StandAloneProject")
            Case "ClientServer"
                rpa_type = [mod].GetType("Rpa00.ClientServerProject")
            Case Else
                Console.WriteLine($"プロジェクト型 '{txt}' は想定されていません")
                Console.ReadLine()
                Exit Sub
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

        If args(2) = "auto" Then
            trn.CommandText = "load"
            Call trn.CreateCommand()
            Call sys.Main(trn, rpa)
        End If

        Do Until trn.ExitFlag
            Console.Write($"{rpa.ProjectAlias} > ")
            trn.CommandText = Console.ReadLine()
            Call trn.CreateCommand()
            Call sys.Main(trn, rpa)
            Console.WriteLine(vbNullString)
        Loop
    End Sub
End Module
