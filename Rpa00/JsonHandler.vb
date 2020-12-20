Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Text

Public Class JsonHandler(Of T As New)
    Private Const SHIFT_JIS As String = "Shift-JIS"
    Private Const UTF8 As String = "utf-8"
    Public Shared MyEncoding As String = UTF8

    'Private Property ModelFileName As String

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

        'Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)

        Load = Nothing

        Try
            sr = New System.IO.StreamReader(
                f, System.Text.Encoding.GetEncoding(MyEncoding))
            txt = sr.ReadToEnd()

            Load = JsonConvert.DeserializeObject(Of T)(txt)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
        End Try
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
                f, System.Text.Encoding.GetEncoding(MyEncoding))
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

    Public Overloads Sub Save(ByVal f As String,
                                   ByVal m As T)
        Dim txt As String
        Dim sw As System.IO.StreamWriter
        'Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        Try
            sw = New System.IO.StreamWriter(
                f, False, System.Text.Encoding.GetEncoding(MyEncoding)
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

    Public Overloads Sub Save(Of T2 As New)(ByVal f As String,
                                          ByVal m As T2)
        Dim txt As String
        Dim sw As System.IO.StreamWriter

        'Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)

        Try
            sw = New System.IO.StreamWriter(
                f, False, System.Text.Encoding.GetEncoding(MyEncoding)
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
