Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ExportMyCommandsCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        If dat.Initializer.MyCommandDictionary.Count = 0 Then
            Console.WriteLine($"コマンド登録簿に登録がありません")
            Return False
        End If
        Return True
    End Function

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Dim txt As String
        Dim sw As System.IO.StreamWriter
        Dim rtn As Integer = -1
        Try
            sw = New System.IO.StreamWriter(
                RpaInitializer.MyCommandsJsonFileName, False, System.Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING)
            )

            txt = JsonConvert.SerializeObject(dat.Initializer.MyCommandDictionary, Formatting.Indented)

            'ファイルへ保存
            sw.Write(txt)
            Console.WriteLine($"コマンドを '{RpaInitializer.MyCommandsJsonFileName}' へエクスポートしました")
            Console.WriteLine()
            rtn = 0
        Catch ex As Exception
            Console.WriteLine($"{ex.Message}")
            Console.WriteLine($"コマンドのエクスポートに失敗しました")
            Console.WriteLine()
            rtn = 1000
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
        End Try
        Return rtn
    End Function
End Class
