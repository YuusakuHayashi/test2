Public Class SetupRpaCommand : Inherits RpaCommandBase
    Private ReadOnly Property ExecuteHandler(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Object
        Get
            Dim cmd As Object
            Select Case trn.MainCommand
                Case "newproject" : cmd = New NewProjectCommand
                Case "setupinitializer" : cmd = New SetupInitializerCommand
                Case "exit" : cmd = New ExitCommand                          ' 現状、SetupのExit と ApplicationのExit に差はないため再利用している
                Case Else : cmd = Nothing
            End Select
            Return cmd
        End Get
    End Property
    '---------------------------------------------------------------------------------------------'

    Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
        Get
            Return {0, 99}
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableUserLevel As Integer
        Get
            Return RpaCodes.ProjectUserLevel.User
        End Get
    End Property

    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        Console.WriteLine($"設定モード移行")
        trn.Modes.Add("setup")
        Do
            trn.CommandText = trn.ShowRpaIndicator(rpa)
            Call trn.CreateCommand()
            Call CommandExecute(trn, rpa, ini)
            Console.WriteLine()
        Loop Until trn.ExitFlag

        trn.ExitFlag = False
        trn.Modes.Remove(trn.Modes.Remove("setup"))
        Console.WriteLine($"設定モード終了")
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
