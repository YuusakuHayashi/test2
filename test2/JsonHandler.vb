Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class JsonHandler(Of T)
    Private Const SHIFT_JIS As String = "Shift-JIS"

    Public Delegate Function LoadProxy(ByVal f As String) As T
    Public Delegate Function LoadProxy(Of T2)(ByVal f As String) As T2
    Public Delegate Sub SaveProxy(ByVal f As String, ByVal m As T)
    Public Delegate Sub SaveProxy(Of T2)(ByVal f As String, ByVal m As T2)

    'Public Function CheckLoadModel(Of T)(ByVal f As String) As Integer
    '    Dim returnCode As Integer : returnCode = -1
    '    Dim m As T
    '    Try
    '        m = Me.ModelLoad(Of T)(f)
    '        returnCode = 0
    '    Catch ex As Exception
    '        returnCode = 99
    '    End Try

    '    CheckLoadModel = returnCode
    'End Function

    Public Overloads Function CheckModel(Of T2)(ByVal f As String) As Boolean
        Dim b As Boolean : b = False
        Dim m As T2
        Try
            m = Me.ModelLoad(Of T2)(f)
            b = True
        Catch ex As Exception
            b = False
        End Try
        CheckModel = b
    End Function

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

    Public Sub NoSave(ByVal f As String, ByVal m As T)
        ' Nothing To Do
    End Sub

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
