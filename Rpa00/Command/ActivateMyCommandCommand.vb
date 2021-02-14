Imports Rpa00

Public Class ActivateMyCommandCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        For Each cmd In dat.Initializer.MyCommandDictionary.Values
            If dat.System.CommandData.Parameters.Contains(cmd.TrueCommand) Then
                cmd.IsEnabled = True
                Console.WriteLine($"'{cmd.TrueCommand}' コマンドを有効にしました")
            End If
        Next
        Console.WriteLine()

        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecutableParameterCount = {1, 999}
        Me.ExecuteIfNoProject = True
    End Sub
End Class
