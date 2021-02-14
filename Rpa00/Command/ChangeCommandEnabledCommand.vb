
Public Class ChangeCommandEnabledCommand : Inherits RpaCommandBase
    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        ' 論理和
        If Not (Boolean.TryParse(dat.System.CommandData.Parameters(1), True) Or Boolean.TryParse(dat.System.CommandData.Parameters(1), False)) Then
            Console.WriteLine($"パラメタ '{dat.System.CommandData.Parameters(1)}' の論理値への変換に失敗しました")
            Console.WriteLine()
            Return False
        End If
        Return True
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim val As String = dat.System.CommandData.Parameters(0)
        Dim torf As Boolean = Boolean.Parse(dat.System.CommandData.Parameters(1))
        Dim hit As Boolean = False

        For Each cmd In dat.Initializer.MyCommandDictionary.Values
            If cmd.TrueCommand = val Then
                cmd.IsEnabled = torf
                hit = True
                Exit For
            End If
        Next
        If Not hit Then
            dat.Initializer.MyCommandDictionary.Add(
                val, (New RpaInitializer.MyCommand With {
                    .TrueCommand = val,
                    .IsEnabled = torf
                })
            )
        End If

        Console.WriteLine($"コマンド '{[val]}' を '{torf.ToString}' に変更しました")
        Console.WriteLine()

        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecutableParameterCount = {2, 2}
        Me.ExecuteIfNoProject = True
    End Sub
End Class
