Imports System
Imports System.IO
Imports System.Reflection

Module Program
    Sub Main(args As String())
        Dim asm As Assembly
        Dim [mod] As [Module]
        Dim rpa_type As Type : Dim trn_type As Type : Dim sys_type As Type
        Dim rpa, trn, sys

        asm = Assembly.LoadFrom("C:\Users\yuusa\project\test2\Rpa00\obj\Debug\Rpa00.dll")
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

        rpa = rpa.ModelLoad(rpa.SystemJsonFileName)
        If rpa Is Nothing Then
            Console.WriteLine(rpa.SYSTEM_JSON_FILENAME & " ÇÃì«Ç›çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩ")
            Console.WriteLine("RPAÇÃé¿çsèIóπÇµÇ‹ÇµÇΩ")
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
