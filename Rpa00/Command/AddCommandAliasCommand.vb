Public Class AddCommandAliasCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim [key] As String = dat.Transaction.Parameters(0)
        Dim [alias] As String = dat.Transaction.Parameters(1)
        If dat.Initializer.MyCommandDictionary.ContainsKey([key]) Then
            dat.Initializer.MyCommandDictionary([key]).[Alias] = [alias]
        Else
            dat.Initializer.MyCommandDictionary.Add(
                [key], New RpaInitializer.MyCommand With {
                    .[Alias] = [alias],
                    .IsEnabled = True
                }
            )
        End If

        Console.WriteLine($"コマンド '{[key]}' に '{[alias]}' を追加しました")
        Console.WriteLine()

        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecutableParameterCount = {2, 2}
        Me.ExecuteIfNoProject = True
    End Sub
End Class
