Imports System
Imports System.IO

Module Program
    Sub Main(args As String())
        Dim trn = New RpaTransaction
        Dim rpa = New RpaProject

        rpa = rpa.ModelLoad(rpa.SystemJsonFileName)
        If rpa Is Nothing Then
            Console.WriteLine(RpaProject.SYSTEM_JSON_FILENAME & " の読み込みに失敗しました")
            Console.WriteLine("RPAの実行終了しました")
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
