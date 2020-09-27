Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class ProjectModel(Of T) : Inherits FileContentsModel
    Private Const SHIFT_JIS = "SHIFT_JIS"

    Public Function DirectoryCheck(ByVal d As String) As Boolean
        d = _ProjectDirectory(d)
        If Directory.Exists(d) Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Function FileCheck(ByVal f As String) As Boolean
        f = ProjectFile(f)
        If File.Exists(f) Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Sub DirectoryEstablish(ByVal d As String)
        d = _ProjectDirectory(d)
        System.IO.Directory.CreateDirectory(d)
    End Sub


    Public Sub FileEstablish(ByVal f As String)
        f = ProjectFile(f)
        System.IO.File.Create(f)
    End Sub


    Public Sub FileWrite(ByVal f As String, ByVal txt As String)
        f = ProjectFile(f)
        Dim sw As StreamWriter
        sw = New System.IO.StreamWriter(
            f, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
        )
        sw.Write(txt)
        sw.Close()
        sw.Dispose()
    End Sub


    Private Property _SaveState As Integer
    Protected Property SaveState As Integer
        Get
            Return _SaveState
        End Get
        Set(value As Integer)
            _SaveState = value
        End Set
    End Property


    Public Delegate Sub LoadedMethodProxy(ByRef m As T)

    Public Overloads Function ModelLoad(ByVal f As String) As T
        Dim lambda As LoadedMethodProxy

        lambda = Sub()
                     ' Nothing To Do
                 End Sub

        ModelLoad = ModelLoad(f, lambda)
    End Function


    Public Overloads Function ModelLoad(ByVal f As String, ByRef lm As LoadedMethodProxy) As T
        Dim txt As String
        Dim sr As System.IO.StreamReader
        Dim m As T

        m = Nothing

        Try
            f = ProjectFile(f)

            sr = New System.IO.StreamReader(
                f, System.Text.Encoding.GetEncoding(SHIFT_JIS))
            txt = sr.ReadToEnd()

            m = JsonConvert.DeserializeObject(Of T)(txt)

            ' ロード後実行したいメソッド
            lm(m)
        Catch ex As Exception
            '
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If

            ModelLoad = m
        End Try
    End Function

    Public Overloads Function ModelLoad(Of T2)(ByVal f As String) As T2
        Dim txt As String
        Dim sr As System.IO.StreamReader

        ModelLoad = Nothing

        Try
            f = ProjectFile(f)

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

    'Public Overloads Function ModelLoad(ByVal f As String) As T
    '    Dim txt As String
    '    Dim sr As System.IO.StreamReader

    '    ModelLoad = Nothing

    '    Try
    '        f = ProjectFile(f)

    '        sr = New System.IO.StreamReader(
    '            f, System.Text.Encoding.GetEncoding(SHIFT_JIS))
    '        txt = sr.ReadToEnd()

    '        ModelLoad = JsonConvert.DeserializeObject(Of T)(txt)
    '    Catch ex As Exception
    '        '
    '    Finally
    '        If sr IsNot Nothing Then
    '            sr.Close()
    '            sr.Dispose()
    '        End If
    '    End Try
    'End Function



    'Public Overloads Sub ModelSave(Of T)(ByVal f As String,
    '                                     ByVal m As T)
    Public Overloads Sub ModelSave(ByVal f As String,
                                   ByVal m As T)
        Dim txt As String
        Dim sw As System.IO.StreamWriter

        Try
            f = ProjectFile(f)

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
            f = ProjectFile(f)

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
                                          ByVal m As T,
                                          ByVal key As String)

        ' T  ... 更新するメンバー
        ' T2 ... 全体のクラス
        Dim txt As String
        Dim txt2 As String
        Dim sw As System.IO.StreamWriter
        Dim sr As System.IO.StreamReader
        Dim obj As T
        Dim test As T

        Try
            f = ProjectFile(f)

            If File.Exists(f) Then
                ' ファイルあり
                sr = New System.IO.StreamReader(
                    f, System.Text.Encoding.GetEncoding(SHIFT_JIS)
                )
                txt = sr.ReadToEnd()
                obj = JsonConvert.DeserializeObject(txt)

                'obj(key) = m

                txt2 = JsonConvert.SerializeObject(obj)
            Else
                ' ファイルなし
                Throw New Exception
            End If

            sw = New System.IO.StreamWriter(
                f, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
            )

            'ファイルへ保存
            sw.Write(txt2)
        Catch ex As Exception
            '
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
        End Try
    End Sub

    'Public Overrides Sub MemberCheck()
    '    ' メンバのデフォルト値を設定したい場合、ここに記述
    '    If String.IsNullOrEmpty(Me.CurrentModelJson) Then
    '        Me.CurrentModelJson = Me.DefaultModelJson
    '    End If
    '    If String.IsNullOrEmpty(Me.CurrentInitViewModelJson) Then
    '        Me.CurrentInitViewModelJson = Me.DefaultInitViewModelJson
    '    End If
    '    If String.IsNullOrEmpty(Me.CurrentDBExplorerViewModelJson) Then
    '        Me.CurrentDBExplorerViewModelJson = Me.DefaultDBExplorerViewModelJson
    '    End If
    'End Sub


    Sub New()

        'Dim mkdir As ProjectProxy : mkdir = AddressOf Me.DirectoryEstablish
        'Dim mkfil As ProjectProxy : mkfil = AddressOf Me.FileEstablish

        'Try
        '    If Not Directory.Exists(Me.ProjectDirectoryName) Then
        '        Call mkdir(Me.ProjectDirectoryName)
        '        Call mkfil(Me._ProjectIniFile)
        '        Call mkfil(Me.FileManagerJson)
        '    Else
        '        If Not File.Exists(Me._ProjectIniFile) Then
        '            Throw New Exception
        '        Else
        '            If Not File.Exists(Me.FileManagerJson) Then
        '                Call mkfil(Me.FileManagerJson)
        '            End If
        '        End If
        '    End If
        'Catch ex As Exception
        '    '
        'Finally
        '    '
        'End Try
    End Sub
End Class
