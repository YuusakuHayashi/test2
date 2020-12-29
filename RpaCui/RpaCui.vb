Imports System.Reflection
Imports System.IO
Imports Rpa00

Module RpaCui
    Public Sub Main()
        Dim ptype As String = vbNullString
        Dim txt As String = vbNullString
        Dim inv As Boolean = False

        If Not Directory.Exists(CommonProject.SystemDirectory) Then
            Directory.CreateDirectory(CommonProject.SystemDirectory)
        End If
        If Not Directory.Exists(CommonProject.SystemDllDirectory) Then
            Directory.CreateDirectory(CommonProject.SystemDllDirectory)
        End If
        If Not File.Exists(CommonProject.System00DllFileName) Then
            Console.WriteLine($"ファイル     '{CommonProject.System00DllFileName}' が存在しません")
            Console.ReadLine()
            Exit Sub
        End If

        Dim asm As Assembly = Nothing
        Dim [mod] As [Module] = Nothing
        'asm = Assembly.LoadFrom(CommonProject.System00DllFileName)
        'asm = Assembly.LoadFrom("C:\Users\yuusa\project\test2\Rpa00\obj\Debug\Rpa00.dll")
        asm = Assembly.LoadFrom("\\Coral\個人情報-林祐\project\wpf\test2\Rpa00\bin\Debug\Rpa00.dll")
        [mod] = asm.GetModule("Rpa00.dll")
        Dim trn_type = [mod].GetType("Rpa00.RpaTransaction")
        Dim sys_type = [mod].GetType("Rpa00.RpaSystem")
        Dim ini_type = [mod].GetType("Rpa00.RpaInitializer")
        Dim rpa As Object = Nothing
        Dim trn = Activator.CreateInstance(trn_type)
        Dim sys = Activator.CreateInstance(sys_type)
        Dim ini = Activator.CreateInstance(ini_type)

        If Not File.Exists(CommonProject.SystemIniFileName) Then
            ini.Save(CommonProject.SystemIniFileName, ini)
        End If
        ini = ini.Load(CommonProject.SystemIniFileName)

        If ini.AutoLoad Then
            If ini.CurrentSolution IsNot Nothing Then
                rpa = RpaModule.LoadCurrentRpa(ini)
            End If
        End If

        'Call rpa.HelloProject()
        'Call rpa.CheckProject()

        ' 実行
        Do Until trn.ExitFlag
            trn.CommandText = trn.ShowRpaIndicator(rpa)
            Call trn.CreateCommand()
            Call sys.Main(trn, rpa, ini)
        Loop
    End Sub
End Module
