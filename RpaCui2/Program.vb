Imports System
Imports System.IO

Module Program
    Sub Main(args As String())
        Dim trn = New RpaTransaction
        Dim rpa = New RpaProject

        rpa = rpa.ModelLoad(rpa.SystemJsonFileName)
        If rpa Is Nothing Then
            Console.WriteLine(RpaProject.SYSTEM_JSON_FILENAME & " ÇÃì«Ç›çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩ")
            Console.WriteLine("RPAÇÃé¿çsèIóπÇµÇ‹ÇµÇΩ")
            Console.ReadLine()
            Exit Sub
        End If

        Do Until trn.ExitFlag
            trn.CommandText = Console.ReadLine()
            Call trn.CreateCommand()
            Call RpaSystem.Main(trn, rpa)
        Loop
    End Sub
End Module
