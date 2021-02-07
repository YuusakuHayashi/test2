Imports Rpa00

Public Class RemoveMyCommandCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        'For Each [key] In dat.Transaction.Parameters
        '    If dat.Initializer.MyCommandDictionary.ContainsKey([key]) Then
        '        dat.Initializer.MyCommandDictionary([key]).IsEnabled = False
        '    Else
        '        dat.Initializer.MyCommandDictionary.Add(
        '            [key], New RpaInitializer.MyCommand With {
        '                .[Alias] = [key],
        '                .IsEnabled = False
        '            }
        '        )
        '    End If
        'Next
        For Each cmd In dat.Initializer.MyCommandDictionary.Values
            If dat.Transaction.Parameters.Contains(cmd.TrueCommand) Then
                cmd.IsEnabled = False
                Console.WriteLine($"'{cmd.TrueCommand}' コマンドを無効にしました")
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
