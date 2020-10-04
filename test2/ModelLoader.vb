Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ModelLoader(Of T)
    Private Const SHIFT_JIS As String = "Shift-JIS"

    ' この関数は呼び出したオブジェクトのモデルをロードします
    Public Overloads Function ModelLoad(ByVal f As String) As T
        Dim txt As String
        Dim sr As StreamReader

        ModelLoad = Nothing

        Try
            sr = New System.IO.StreamReader(
                f, System.Text.Encoding.GetEncoding(SHIFT_JIS))
            txt = sr.ReadToEnd()

            ModelLoad = JsonConvert.DeserializeObject(Of T)(txt)
        Catch ex As Exception
            '
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
        End Try
    End Function

    ' この関数は指定したジェネリックモデルをロードします
    Public Overloads Function ModelLoad(Of T2)(ByVal f As String) As T2
        Dim txt As String
        Dim sr As IO.StreamReader

        ModelLoad = Nothing
        Try
            sr = New System.IO.StreamReader(
                f, System.Text.Encoding.GetEncoding(SHIFT_JIS))
            txt = sr.ReadToEnd()

            ModelLoad = JsonConvert.DeserializeObject(Of T2)(txt)
        Catch ex As Exception
            '
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
        End Try
    End Function


    Public Overloads Sub ModelSave(ByVal f As String,
                                   ByVal m As T)
        Dim txt As String
        Dim sw As System.IO.StreamWriter

        Try
            sw = New System.IO.StreamWriter(
                f, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
            )

            txt = JsonConvert.SerializeObject(m, Formatting.Indented)

            'ファイルへ保存
            sw.Write(txt)
        Catch ex As Exception
            '
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
        End Try
    End Sub

    Public Overloads Sub ModelSave(Of T2)(ByVal f As String,
                                          ByVal m As T2)
        Dim txt As String
        Dim sw As System.IO.StreamWriter

        Try
            sw = New System.IO.StreamWriter(
                f, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
            )

            txt = JsonConvert.SerializeObject(m, Formatting.Indented)

            'ファイルへ保存
            sw.Write(txt)
        Catch ex As Exception
            '
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
        End Try
    End Sub
End Class
