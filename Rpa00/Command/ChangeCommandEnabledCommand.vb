﻿
Public Class ChangeCommandEnabledCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        ' 論理和
        If Not (Boolean.TryParse(dat.Transaction.Parameters(1), True) Or Boolean.TryParse(dat.Transaction.Parameters(1), False)) Then
            Console.WriteLine($"パラメタ '{dat.Transaction.Parameters(1)}' の論理値への変換に失敗しました")
            Console.WriteLine()
            Return False
        End If
        Return True
    End Function

    Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
        Get
            Return {2, 2}
        End Get
    End Property

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim [key] As String = dat.Transaction.Parameters(0)
        Dim torf As Boolean = Boolean.Parse(dat.Transaction.Parameters(1))
        If dat.Initializer.MyCommandDictionary.ContainsKey([key]) Then
            dat.Initializer.MyCommandDictionary([key]).IsEnabled = torf
        Else
            dat.Initializer.MyCommandDictionary.Add(
                [key], New RpaInitializer.MyCommand With {
                    .[Alias] = [key],
                    .IsEnabled = torf
                }
            )
        End If

        Console.WriteLine($"コマンド '{[key]}' を '{torf.ToString}' に変更しました")
        Console.WriteLine()

        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
    End Sub
End Class
