Public Class AddCommandAliasCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim [key] As String = dat.System.CommandData.Parameters(1)
        Dim truecmd As String = dat.System.CommandData.Parameters(0)

        If dat.Initializer.MyCommandDictionary.ContainsKey([key]) Then
            dat.Initializer.MyCommandDictionary([key]).TrueCommand = truecmd
        Else
            dat.Initializer.MyCommandDictionary.Add(
                [key], New RpaInitializer.MyCommand With {
                    .TrueCommand = truecmd,
                    .IsEnabled = True
                }
            )
        End If

        Console.WriteLine($"コマンド '{[key]}' に '{truecmd}' を追加しました")
        Console.WriteLine()

        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecutableParameterCount = {2, 2}
        Me.ExecuteIfNoProject = True
    End Sub
End Class
