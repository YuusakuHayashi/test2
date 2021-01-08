Imports System.IO
Imports Rpa00

Public Class Rpa07 : Inherits RpaBase(Of Rpa07)

    Private _PatternFile As String
    Public Property PatternFile As String
        Get
            Return Me._PatternFile
        End Get
        Set(value As String)
            Me._PatternFile = value
        End Set
    End Property


    Private _PatternsDirectory As String
    Public Property PatternsDirectory As String
        Get
            Return Me._PatternsDirectory
        End Get
        Set(value As String)
            Me._PatternsDirectory = value
        End Set
    End Property


    Public Overrides Function CanExecute(ByRef dat As Object) As Boolean
        Return True
    End Function

    Public Overrides Function Execute(ByRef dat As Object) As Integer
        Do
            Console.WriteLine($"処理を選択してください")
            Console.WriteLine($"1...パターンファイルを更新  2...パターン定義ファイルの確認  9...終了")
            Dim idxtext As String = vbNullString
            Do
                idxtext = vbNullString
                idxtext = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
            Loop Until idxtext = "1" Or idxtext = "2" Or idxtext = "9"

            If idxtext = "9" Then
                Exit Do
            End If

            Dim i As Integer = -1
            If idxtext = "1" Then
                i = UpdatePattern(dat)
            ElseIf idxtext = "2" Then
                i = ConfirmPattern(dat)
            End If

            Dim yorn As String = vbNullString
        Loop Until False

        Return 0
    End Function

    Private Function UpdatePattern(ByRef dat As Object) As Integer
        Dim ptns As List(Of String) = Directory.GetFiles(Me.PatternsDirectory).ToList
        For Each ptn In ptns
            Console.WriteLine($"{ptn.IndexOf(ptn)}  {Path.GetFileName(ptn)}")
        Next

        Dim idx As Integer = -1
        Do
            Console.WriteLine()
            Console.WriteLine($"インデックス番号を選択してください")
            Dim idxtext As String = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()

            If Not IsNumeric(idxtext) Then
                Console.WriteLine($"インデックス番号が不正です")
                Continue Do
            End If

            idx = -1
            idx = Integer.Parse(idxtext)

            If idx < 0 Or idx > (ptns.Count - 1) Then
                Console.WriteLine($"インデックス番号が不正です")
                Continue Do
            End If

            Dim yorn As String = vbNullString
            Do
                yorn = vbNullString
                Console.WriteLine($"よろしいですか？ {idx}  {Path.GetFileName(ptns(idx))}")
                yorn = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn = "y" Or yorn = "n" Then
                    Exit Do
                End If
            Loop Until False

            If yorn = "y" Then
                Exit Do
            End If
        Loop Until False

        File.Copy(ptns(idx), Me.PatternFile, True)
        Console.WriteLine($"{Path.GetFileName(Me.PatternFile)} <- {Path.GetFileName(ptns(idx))}")
        Console.WriteLine()

        Return 0
    End Function

    Private Function ConfirmPattern(ByRef dat As Object) As Integer
        Return 0
    End Function

    Public Overrides Function SetupProjectObject(project As String) As Object
        Throw New NotImplementedException()
    End Function
End Class
