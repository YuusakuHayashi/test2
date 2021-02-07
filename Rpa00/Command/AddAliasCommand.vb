Public Class AddAliasCommand : Inherits RpaCommandBase
    Private _AliasText As String
    Private Property AliasText As String
        Get
            Return Me._AliasText
        End Get
        Set(value As String)
            Me._AliasText = value
        End Set
    End Property

    Private _TrueCommandText As String
    Private Property TrueCommandText As String
        Get
            Return Me._TrueCommandText
        End Get
        Set(value As String)
            Me._TrueCommandText = value
        End Set
    End Property

    Private Function Check(ByRef dat As RpaDataWrapper) As Integer
        Console.WriteLine()

        Console.WriteLine($"別名コマンドを入力してください")
        Me.AliasText = dat.Transaction.ShowRpaIndicator(dat)
        Console.WriteLine()

        Console.WriteLine($"実行コマンドを入力してください")
        Me.TrueCommandText = dat.Transaction.ShowRpaIndicator(dat)
        Console.WriteLine()

        Console.WriteLine($"別名コマンド '{Me.AliasText}'")
        Console.WriteLine($"実行コマンド '{Me.TrueCommandText}'")
        Do
            Console.WriteLine($"よろしいですか？ (y/n)")
            Dim yorn As String = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()

            If yorn = "y" Then
                Return True
            End If
            If yorn = "n" Then
                Return False
            End If
        Loop Until False

        Return 0
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        If dat.Initializer.AliasDictionary.ContainsKey(Me.AliasText) Then
            dat.Initializer.AliasDictionary(Me.AliasText) = Me.TrueCommandText
        Else
            dat.Initializer.AliasDictionary.Add(Me.AliasText, Me.TrueCommandText)
        End If
        Return 0
    End Function

    Sub New()
        Me.ExecuteIfNoProject = True
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
