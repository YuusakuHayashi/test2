Public Class ShowMyCommandCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim aliascmdlen As Integer = 20
        Dim truecmdlen As Integer = 20
        Console.WriteLine()
        Console.WriteLine($" Alias Command        | True Command         | IsEnabled")
        Console.WriteLine($"________________________________________________________")

        For Each cmd In dat.Initializer.MyCommandDictionary
            Dim aliascmd As String = cmd.Key
            If aliascmd.Length > aliascmdlen Then
                aliascmd = Strings.Left(aliascmd, (aliascmdlen - 3)) & "..."
            ElseIf aliascmd.Length < aliascmdlen Then
                aliascmd = aliascmd & Strings.StrDup((aliascmdlen - aliascmd.Length), " "c)
            Else
                'Nothing To Do
            End If

            Dim truecmd As String = cmd.Value.TrueCommand
            If truecmd.Length > truecmdlen Then
                truecmd = Strings.Left(truecmd, (truecmdlen - 3)) & "..."
            ElseIf truecmd.Length < truecmdlen Then
                truecmd = truecmd & Strings.StrDup((truecmdlen - truecmd.Length), " "c)
            Else
                'Nothing To Do
            End If

            Console.WriteLine($" {aliascmd} | {truecmd} | {cmd.Value.IsEnabled.ToString}")
        Next

        If dat.Initializer.MyCommandDictionary.Count = 0 Then
            Console.WriteLine($"コマンドの登録はありません")
        End If
        Console.WriteLine($"________________________________________________________")
        Console.WriteLine()
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecuteIfNoProject = True
    End Sub
End Class
