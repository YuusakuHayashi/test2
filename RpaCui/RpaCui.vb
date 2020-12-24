Imports System.Reflection
Imports System.IO
Imports Rpa00

Module RpaCui
    Public Sub Main()
        Dim ptype As String = vbNullString
        Dim txt As String = vbNullString
        Dim inv As Boolean = False
        Dim arch As Integer

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
        asm = Assembly.LoadFrom("C:\Users\yuusa\project\test2\Rpa00\obj\Debug\Rpa00.dll")
        [mod] = asm.GetModule("Rpa00.dll")
        Dim trn_type = [mod].GetType("Rpa00.RpaTransaction")
        Dim sys_type = [mod].GetType("Rpa00.RpaSystem")
        Dim ini_type = [mod].GetType("Rpa00.RpaInitializer")
        Dim rpa As Object = Nothing
        Dim trn = Activator.CreateInstance(trn_type)
        Dim sys = Activator.CreateInstance(sys_type)
        Dim ini = Activator.CreateInstance(ini_type)
        Dim ics_type = [mod].GetType("Rpa00.IntranetClientServerProject")
        Dim sap_type = [mod].GetType("Rpa00.StandAloneProject")
        Dim csp_type = [mod].GetType("Rpa00.ClientServerProject")
        Dim ics = Activator.CreateInstance(ics_type)
        Dim sap = Activator.CreateInstance(sap_type)
        Dim csp = Activator.CreateInstance(csp_type)

        If Not File.Exists(CommonProject.SystemIniFileName) Then
            ini.Save(CommonProject.SystemIniFileName, ini)
        End If
        ini = ini.Load(CommonProject.SystemIniFileName)
        If ini.CurrentSolution Is Nothing Then
            rpa = Nothing
        Else
            ' 動的オブジェクト生成
            Select Case ini.CurrentSolution.Architecture
                Case ics.SystemArchType : rpa = ics
                Case sap.SystemArchType : rpa = sap
                Case csp.SystemArchType : rpa = csp
            End Select
        End If

        'Call rpa.HelloProject()
        'Call rpa.CheckProject()

        ' 実行
        Do Until trn.ExitFlag
            If rpa IsNot Nothing Then
                Console.Write($"{rpa.ProjectAlias}>")
            Else
                Console.Write("NoRpa>")
            End If
            trn.CommandText = Console.ReadLine()
            Call trn.CreateCommand()
            Call sys.Main(trn, rpa, ini)
            Console.WriteLine(vbNullString)
        Loop
    End Sub

End Module
