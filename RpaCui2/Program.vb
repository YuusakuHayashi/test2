Imports System

Module Program
    Sub Main(args As String())
        Dim trn = New RpaTransaction
        Dim rpa = New RpaProject
        Dim rc = -1

        rpa = rpa.ModelLoad(rpa.SystemJsonFileName)
        If rpa Is Nothing Then
            rpa = New RpaProject
            rpa.ModelSave(rpa.SystemJsonFileName, rpa)
        End If

        'Call rpa.MakeStaff()
        'Call rpa.CheckSystemConstitution()

        Do Until trn.ExitFlag
            trn.CommandText = Console.ReadLine()
            Call trn.CreateCommand()
            Call RpaSystem.Main(trn, rpa)
        Loop
    End Sub
End Module
