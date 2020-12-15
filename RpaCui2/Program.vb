Imports System
Imports System.IO
Imports System.Reflection

Module Program
    Sub Main(args As String())
        Dim asm As Assembly
        Dim [mod] As [Module]
        Dim rpa_type As Type : Dim trn_type As Type : Dim sys_type As Type
        Dim rpa, trn, sys
        Dim rpadir As String
        Dim dlldir As String
        Dim rpa00dll As String
        Dim json As String

        rpadir = Environment.GetEnvironmentVariable("USERPROFILE") & "\rpa_project"
        dlldir = $"{rpadir}\dll"
        If args.Count > 1 Then
            If args(0) = "Debug" Then
                dlldir = args(1)
            End If
        End If
        rpa00dll = $"{dlldir}\Rpa00.dll"

        If Not Directory.Exists(rpadir) Then
            Directory.CreateDirectory(rpadir)
        End If
        If Not Directory.Exists(dlldir) Then
            Directory.CreateDirectory(dlldir)
        End If
        If Not File.Exists(rpa00dll) Then
            Console.WriteLine($"DLL '{rpa00dll}' ������܂���B���L�҂�����肵�Ă�������   DATE-WRITTEN: 2020/12/14")
            Console.ReadLine()
            Exit Sub
        End If

        asm = Assembly.LoadFrom(rpa00dll)
        [mod] = asm.GetModule("Rpa00.dll")
        rpa_type = [mod].GetType("Rpa00.RpaProject")
        trn_type = [mod].GetType("Rpa00.RpaTransaction")
        sys_type = [mod].GetType("Rpa00.RpaSystem")

        If rpa_type IsNot Nothing Then
            rpa = Activator.CreateInstance(rpa_type)
        End If
        If trn_type IsNot Nothing Then
            trn = Activator.CreateInstance(trn_type)
        End If
        If sys_type IsNot Nothing Then
            sys = Activator.CreateInstance(sys_type)
        End If

        ' rpa.SystemJsonFileName ���Q�Ƃ���ۂɁA�Ȃ���΃t�@�C�������������B
        json = rpa.SystemJsonFileName
        rpa = rpa.ModelLoad(json)
        If rpa Is Nothing Then
            Console.WriteLine($"�ݒ�t�@�C�� '{json}' �̓ǂݍ��݂Ɏ��s���܂���")
            Console.ReadLine()
            Exit Sub
        End If

        Do Until trn.ExitFlag
            trn.CommandText = Console.ReadLine()
            Call trn.CreateCommand()
            Call sys.Main(trn, rpa)
        Loop
    End Sub
End Module
