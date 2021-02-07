Imports System.IO

Public Class SelectMyDirectoryCommand : Inherits RpaCommandBase
    Private _NewMyDirectory As String
    Private Property NewMyDirectory As String
        Get
            Return Me._NewMyDirectory
        End Get
        Set(value As String)
            Me._NewMyDirectory = value
        End Set
    End Property

    Private _NewMyDirectories As List(Of String)
    Private Property NewMyDirectories As List(Of String)
        Get
            If Me._NewMyDirectories Is Nothing Then
                Me._NewMyDirectories = New List(Of String)
            End If
            Return Me._NewMyDirectories
        End Get
        Set(value As List(Of String))
            Me._NewMyDirectories = value
        End Set
    End Property

    Private Function Check(ByRef dat As RpaDataWrapper) As Integer
        If dat.Project.MyDirectories.Count <= 0 Then
            Console.WriteLine($"'MyDirectories' にリストが存在しません")
            Return False
        End If

        Me.NewMyDirectories = dat.Project.MyDirectories

        Do
            Dim idx As Integer = -1
            Console.WriteLine($" ID  | ディレクトリ名")
            Console.WriteLine($"-----+----------------------------------------------------------------")
            For Each mydir In Me.NewMyDirectories
                Dim idxno As String = String.Format("{0, 4}", Me.NewMyDirectories.IndexOf(mydir))
                Console.WriteLine($"{idxno} | {mydir}")
            Next
            Console.WriteLine()
            Do
                idx = -1
                Console.WriteLine("'ID'を選択してください")
                Dim idxstr As String = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()

                If Not IsNumeric(idxstr) Then
                    Continue Do
                End If

                idx = Integer.Parse(idxstr)
                If idx < Me.NewMyDirectories.IndexOf(Me.NewMyDirectories.First) Then
                    Continue Do
                End If
                If idx > Me.NewMyDirectories.IndexOf(Me.NewMyDirectories.Last) Then
                    Continue Do
                End If
                Exit Do
            Loop Until False

            If Not Directory.Exists(Me.NewMyDirectories(idx)) Then
                Console.WriteLine($"ディレクトリ '{Me.NewMyDirectories(idx)}' は存在しません")
                Return False
            End If

            Do
                Console.WriteLine($"よろしいですか？ '{Me.NewMyDirectories(idx)}' (y/n)")
                Dim yorn As String = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn = "y" Then
                    Me.NewMyDirectory = Me.NewMyDirectories(idx)
                    Return True
                End If
                If yorn = "n" Then
                    Exit Do
                End If
            Loop Until False
        Loop Until False
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        dat.Project.MyDirectory = Me.NewMyDirectory
        Console.WriteLine($"ディレクトリ '{Me.NewMyDirectory}' をセットしました")
        Console.WriteLine()
    End Function

    Sub New()
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
