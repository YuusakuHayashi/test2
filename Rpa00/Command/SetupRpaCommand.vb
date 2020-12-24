Public Class SetupRpaCommand : Inherits RpaCommandBase

    Private ReadOnly Property ExecuteHandler(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Object
        Get
            Dim cmd As Object
            Select Case trn.MainCommand
                Case "newproject" : cmd = New NewProjectCommand
                Case Else : cmd = Nothing
            End Select
            Return cmd
        End Get
    End Property
    '---------------------------------------------------------------------------------------------'

    Public Overrides ReadOnly Property ExecutableProjectArchitectures As Integer()
        Get
            Return {
                (New IntranetClientServerProject).SystemArchType,
                (New StandAloneProject).SystemArchType,
                (New ClientServerProject).SystemArchType
            }
        End Get
    End Property

    Public Overrides ReadOnly Property CanExecute(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Boolean
        Get
            Return True
        End Get
    End Property

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
        Do
            Console.Write("setup>")
            trn.CommandText = Console.ReadLine()
            Call trn.CreateCommand()
            Call CommandExecute(trn, rpa, ini)
            Console.WriteLine(vbNullString)
        Loop Until trn.ExitFlag

        trn.ExitFlag = False
        Console.WriteLine($"プロジェクト設定変更を終了します")
        Console.ReadLine()
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
