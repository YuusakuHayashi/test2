Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Text

Public Class JsonHandler(Of T As New)
    Private Const SHIFT_JIS As String = "Shift-JIS"
    Private Const UTF8 As String = "utf-8"

    <JsonIgnore>
    Public Property FirstLoad As Boolean

    ' ロード可能かどうかのチェック
    Public Overloads Function CheckModel(Of T2 As New)(ByVal f As String) As Boolean
        Dim b As Boolean : b = False
        Dim m As T2
        Try
            m = Me.Load(Of T2)(f)
            b = True
        Catch ex As Exception
            b = False
        End Try
        CheckModel = b
    End Function

    Public Overloads Function CheckModel(ByVal f As String) As Boolean
        Return (CheckModel(Of T)(f))
    End Function

    ' この関数は呼び出したオブジェクトのモデルをロードします
    Public Overloads Function Load(ByVal f As String) As T
        Dim txt As String
        Dim sr As StreamReader
        Dim rtn As Object

        Try
            sr = New System.IO.StreamReader(
                f, System.Text.Encoding.GetEncoding(SHIFT_JIS))
            txt = sr.ReadToEnd()

            rtn = JsonConvert.DeserializeObject(Of T)(txt)
            rtn.FirstLoad = True
        Catch ex As Exception
            rtn = Nothing
            Console.WriteLine(ex.Message)
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
        End Try
        Return rtn
    End Function

    ' この関数は指定したジェネリックモデルをロードします
    'Public Overloads Function Load(Of T2 As New)() As T2
    '    Load = Load(Of T2)(Me.ModelFileName)
    'End Function

    ' この関数は指定したジェネリックモデルをロードします
    Public Overloads Function Load(Of T2 As New)(ByVal f As String) As T2
        Dim txt As String
        Dim sr As IO.StreamReader

        'Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)

        Load = Nothing
        Try
            sr = New System.IO.StreamReader(
                f, System.Text.Encoding.GetEncoding(SHIFT_JIS))
            txt = sr.ReadToEnd()

            Load = JsonConvert.DeserializeObject(Of T2)(txt)
        Catch ex As Exception
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

    Public Overloads Function GetSerializedText(ByVal m As T) As String
        Dim txt As String = vbNullString
        Try
            txt = JsonConvert.SerializeObject(m, Formatting.Indented)
        Catch ex As Exception
            txt = vbNullString
        End Try
        Return txt
    End Function

    Public Overloads Sub Save(ByVal f As String, ByVal m As T)
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
            Console.WriteLine(ex.Message)
            Console.WriteLine($"{m.GetType.Name} のセーブに失敗しました")
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
        End Try
    End Sub

    Public Overloads Sub Save(Of T2 As New)(ByVal f As String, ByVal m As T2)
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
