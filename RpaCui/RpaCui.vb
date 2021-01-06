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
        Dim dat_type = [mod].GetType("Rpa00.RpaDataWrapper")
        Dim dat = Activator.CreateInstance(dat_type)

        If Not File.Exists(RpaInitializer.SystemIniFileName) Then
            dat.Initializer.Save(RpaInitializer.SystemIniFileName, dat.Initializer)
        End If
        dat.Initializer = dat.Initializer.Load(RpaInitializer.SystemIniFileName)

        dat.Project = RpaModule.LoadCurrentRpa(dat)

        'Call rpa.HelloProject()
        'Call rpa.CheckProject()

        ' 実行
        Do Until dat.Transaction.ExitFlag
            Call dat.System.Main(dat)
        Loop
    End Sub
End Module
