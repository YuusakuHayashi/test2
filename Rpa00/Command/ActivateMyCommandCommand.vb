Imports Rpa00

Public Class ActivateMyCommandCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        For Each [key] In dat.Transaction.Parameters
            If dat.Initializer.MyCommandDictionary.ContainsKey([key]) Then
                dat.Initializer.MyCommandDictionary([key]).IsEnabled = True
            Else
                dat.Initializer.MyCommandDictionary.Add(
                    [key], New RpaInitializer.MyCommand With {
                        .[Alias] = [key],
                        .IsEnabled = True
                    }
                )
            End If
        Next

        Console.WriteLine($"コマンドを無効にしました")
        Console.WriteLine()

        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecutableParameterCount = {1, 999}
        Me.ExecuteIfNoProject = True
    End Sub
End Class
