Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ImportMyCommandsCommand : Inherits RpaCommandBase
    Private _MyCommandJsonFileName As String

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim rtn As Integer = -1
        Dim sr As StreamReader
        Try
            sr = New System.IO.StreamReader(
                Me._MyCommandJsonFileName, System.Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING))

            Dim txt As String = sr.ReadToEnd()

            Dim dic As Dictionary(Of String, RpaInitializer.MyCommand) _
                = JsonConvert.DeserializeObject(Of Dictionary(Of String, RpaInitializer.MyCommand))(txt)

            dat.Initializer.MyCommandDictionary = dic
            Console.WriteLine($"コマンドをインポートしました")
            Console.WriteLine()
        Catch ex As Exception
            Console.WriteLine($"{ex.Message}")
            Console.WriteLine($"コマンドのインポートに失敗しました")
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

    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        If dat.Initializer.MyCommandDictionary.Count > 0 Then
            Dim yorn As String = vbNullString
            Do
                yorn = vbNullString
                Console.WriteLine()
                Console.WriteLine($"インポートを実行すると、現在のコマンド設定を上書きします")
                Console.WriteLine($"よろしいですか？ (y/n)")
                yorn = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
            Loop Until yorn = "y" Or yorn = "n"
            If yorn = "n" Then
                Return False
            End If
        End If

        Me._MyCommandJsonFileName = vbNullString
        Me._MyCommandJsonFileName = RpaModule.SetFileFromDialog(dat, "MyCommandFile")

        If String.IsNullOrEmpty(Me._MyCommandJsonFileName) Then
            Console.WriteLine("ファイルが指定されていません")
            Console.WriteLine()
            Return False
        End If

        Dim yorn2 As String = vbNullString
        Do
            yorn2 = vbNullString
            Console.WriteLine($"インポート元ファイル '{Me._MyCommandJsonFileName}'")
            Console.WriteLine($"よろしいですか？ (y/n)")
            yorn2 = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
        Loop Until yorn2 = "y" Or yorn2 = "n"
        If yorn2 = "n" Then
            Return False
        End If

        Return True
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecuteIfNoProject = True
    End Sub
End Class
