﻿Imports Rpa00

Public Class ActivateMyCommandCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
        Get
            Return {1, 999}
        End Get
    End Property

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
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
End Class