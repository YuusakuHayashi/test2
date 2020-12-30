Imports System.Reflection

Public Class SetupProjectCommand : Inherits RpaCommandBase
    Private ReadOnly Property ExecuteHandler(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Object
        Get
            Dim cmd As Object
            Select Case trn.MainCommand
                Case "changeproperty" : cmd = New ChangeProjectPropertyCommand
                Case "showproperties" : cmd = New ShowProjectPropertiesCommand
                Case "folderbrowser" : cmd = New ChangeProjectPropertyUsingFolderBrowserCommand
                ' 現状、SetupInitializerのExit と ApplicationのExit に差はないため再利用している
                Case "exit" : cmd = New ExitCommand
                Case Else : cmd = Nothing
            End Select
            Return cmd
        End Get
    End Property
    '---------------------------------------------------------------------------------------------'

    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        trn.Modes.Add("project")

        ' 初回時にプロパティ一覧を表示
        trn.MainCommand = "showproperties"
        Call CommandExecute(trn, rpa, ini)

        Do
            trn.CommandText = trn.ShowRpaIndicator(rpa)
            Call trn.CreateCommand()
            Call CommandExecute(trn, rpa, ini)
        Loop Until trn.ExitFlag

        trn.ExitFlag = False
        trn.Modes.Remove(trn.Modes.Remove("project"))
    End Function

    Public Sub CommandExecute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer)
        Dim cmd = ExecuteHandler(trn, rpa, ini)
        If cmd IsNot Nothing Then
            trn.ReturnCode = cmd.Execute(trn, rpa, ini)
        Else
            Console.WriteLine($"コマンド : '{trn.CommandText}' はありません")
        End If
    End Sub
End Class
