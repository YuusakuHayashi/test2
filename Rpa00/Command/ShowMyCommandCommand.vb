Public Class ShowMyCommandCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim truecmdlen As Integer = 20
        Dim aliascmdlen As Integer = 20
        Console.WriteLine()
        Console.WriteLine($" Command              | Alias Command        | IsEnabled")
        Console.WriteLine($"________________________________________________________")

        For Each cmd In dat.Initializer.MyCommandDictionary
            Dim truecmd As String = cmd.Key
            If truecmd.Length > truecmdlen Then
                truecmd = Strings.Left(truecmd, (truecmdlen - 3)) & "..."
            ElseIf truecmd.Length < truecmdlen Then
                truecmd = truecmd & Strings.StrDup((truecmdlen - truecmd.Length), " "c)
            Else
                'Nothing To Do
            End If

            Dim aliascmd As String = cmd.Value.Alias
            If aliascmd.Length > aliascmdlen Then
                aliascmd = Strings.Left(aliascmd, (aliascmdlen - 3)) & "..."
            ElseIf aliascmd.Length < aliascmdlen Then
                aliascmd = aliascmd & Strings.StrDup((aliascmdlen - aliascmd.Length), " "c)
            Else
                'Nothing To Do
            End If

            Console.WriteLine($" {truecmd} | {aliascmd} | {cmd.Value.IsEnabled.ToString}")
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
    End Sub
End Class
