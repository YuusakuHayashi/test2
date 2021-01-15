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


    Private Function Check(ByRef dat As Object) As Boolean
        If String.IsNullOrEmpty(Me.PatternFile) Then
            Me.PatternFile = dat.Project.RootRobotObject.PatternFile
        End If
        If String.IsNullOrEmpty(Me.PatternFile) Then
            Console.WriteLine($"'PatternsDirectory' が設定されていません")
            Return False
        End If
        If String.IsNullOrEmpty(Me.PatternsDirectory) Then
            Me.PatternsDirectory = dat.Project.RootRobotObject.PatternsDirectory
        End If
        If String.IsNullOrEmpty(Me.PatternsDirectory) Then
            Console.WriteLine($"'PatternsDirectory' が設定されていません")
            Return False
        End If
        If Not File.Exists(Me.PatternFile) Then
            Console.WriteLine($"ファイル 'PatternFile' が存在しません")
            Return False
        End If
        If Not Directory.Exists(Me.PatternsDirectory) Then
            Console.WriteLine($"ディレクトリ '{Me.PatternsDirectory}' が存在しません")
            Return False
        End If
        Return True
    End Function

    Private Function Main(ByRef dat As Object) As Integer
        Do
            Console.WriteLine()
            Console.WriteLine($"処理を選択してください")
            Console.Write($" 1...パターンファイルを更新 ")
            Console.Write($" 2...パターン定義ファイルの確認 ")
            Console.Write($" 9...終了 ")
            Console.WriteLine()
            Dim idxtext As String = vbNullString
            idxtext = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()

            Dim i As Integer
            If idxtext = "1" Then
                i = UpdatePattern(dat)
            ElseIf idxtext = "2" Then
                i = ConfirmPattern(dat)
            ElseIf idxtext = "9" Then
                Console.WriteLine($"処理を終了しました")
                Console.WriteLine()
                Exit Do
            Else
                Console.WriteLine($"選択が不正です --> {idxtext}")
                Console.WriteLine()
            End If
        Loop Until False

        Return 0
    End Function

    'Private Function IntractiveModeExecute(ByRef dat As Object) As Integer
    '    Do
    '        Console.WriteLine($"処理を選択してください")
    '        Console.WriteLine($"1...パターンファイルを更新  2...パターン定義ファイルの確認  9...終了")
    '        Do
    '            Dim idxtext As String = vbNullString
    '            idxtext = dat.Transaction.ShowRpaIndicator(dat)
    '            Console.WriteLine()
    '            'i = IndicatedModeExecute(dat, idxtext)
    '        Loop Until idxtext = "1" Or idxtext = "2" Or idxtext = "9"

    '        Dim i As Integer
    '        If idxtext = "1" Then
    '            i = UpdatePattern(dat)
    '        End If
    '        If idxtext = "2" Then
    '            i = ConfirmPattern(dat)
    '        End If
    '        If idxtext = "9" Then
    '            Exit Do
    '        End If
    '    Loop Until False

    '    Return 0
    'End Function

    'Private Function IndicatedModeExecute(ByRef dat As Object, ByVal mode As String) As Integer
    '    Dim i As Integer = -1
    '    If mode = "1" Then
    '        i = UpdatePattern(dat)
    '        i = 1
    '    ElseIf mode = "2" Then
    '        i = ConfirmPattern(dat)
    '        i = 2
    '    ElseIf mode = "9" Then
    '        Console.WriteLine($"処理を終了しました")
    '        Console.WriteLine()
    '        i = 9
    '    Else
    '        Console.WriteLine($"指定 '{mode}' は不正です")
    '        Console.WriteLine()
    '        i = -1
    '    End If
    '    Return i
    'End Function


    Private Function UpdatePattern(ByRef dat As Object) As Integer
        Dim ptns As List(Of String) = Directory.GetFiles(Me.PatternsDirectory).ToList
        For Each ptn In ptns
            Console.WriteLine($"{ptns.IndexOf(ptn)}  {Path.GetFileName(ptn)}")
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
                Console.WriteLine($"選択したファイル内容を確認しますか？ {idx}  {Path.GetFileName(ptns(idx))} (y/n)")
                yorn = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn = "y" Or yorn = "n" Then
                    Exit Do
                End If
            Loop Until False
            If yorn = "y" Then
                Dim i As Integer = ShowPattern(ptns(idx))
            End If

            Dim yorn2 As String = vbNullString
            Do
                yorn2 = vbNullString
                Console.WriteLine($"よろしいですか？ {idx}  {Path.GetFileName(ptns(idx))} (y/n)")
                yorn2 = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn2 = "y" Or yorn2 = "n" Then
                    Exit Do
                End If
            Loop Until False
            If yorn2 = "y" Then
                Exit Do
            End If
        Loop Until False

        File.Copy(ptns(idx), Me.PatternFile, True)
        Console.WriteLine($"{Path.GetFileName(ptns(idx))} -> {Path.GetFileName(Me.PatternFile)}")
        Console.WriteLine("パターンファイルを更新しました")
        Console.WriteLine()

        Return 0
    End Function

    Private Function ConfirmPattern(ByRef dat As Object) As Integer
        'Dim ptns As List(Of String) = Directory.GetFiles(Me.PatternsDirectory).ToList
        'For Each ptn In ptns
        '    Console.WriteLine($"{ptns.IndexOf(ptn)}  {Path.GetFileName(ptn)}")
        'Next

        'Dim idx As Integer = -1
        Console.WriteLine()
        'Do
        '    Console.WriteLine($"インデックス番号を選択してください")
        '    Dim idxtext As String = dat.Transaction.ShowRpaIndicator(dat)
        '    Console.WriteLine()
        '    If Not IsNumeric(idxtext) Then
        '        Console.WriteLine($"インデックス番号が不正です")
        '        Console.WriteLine()
        '        Continue Do
        '    End If

        '    idx = -1
        '    idx = Integer.Parse(idxtext)

        '    If idx < 0 Or idx > (ptns.Count - 1) Then
        '        Console.WriteLine($"インデックス番号が不正です")
        '        Console.WriteLine()
        '        Continue Do
        '    End If

        '    Exit Do
        'Loop Until False

        Dim i As Integer = ShowPattern(Me.PatternFile)

        Return 0
    End Function

    Private Function ShowPattern(ByVal pfile As String) As Integer
        Dim sr As StreamReader
        Dim txt As String = vbNullString
        Dim rtn = -1
        Try
            sr = New System.IO.StreamReader(
                pfile, System.Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING))
            txt = sr.ReadToEnd.TrimEnd
            Console.WriteLine(txt)
            Console.WriteLine()
            rtn = 0
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine()
            rtn = 1000
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
        End Try
        Return rtn
    End Function

    Public Overrides Function SetupProjectObject(project As String) As Object
        Throw New NotImplementedException()
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
    End Sub
End Class
