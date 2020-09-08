Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ProjectModel(Of T)
    Private Const SHIFT_JIS = "SHIFT_JIS"

    Protected ReadOnly Property ProjectDirectoryName As String
        Get
            Return System.Environment.GetEnvironmentVariable("USERPROFILE") & "\test"
        End Get
    End Property

    Private ReadOnly Property _ProjectIniFileName As String
        Get
            Return Me.ProjectDirectoryName & "\ProjectInit"
        End Get
    End Property

    Public ReadOnly Property FileManagerFileName As String
        Get
            Return Me.ProjectDirectoryName & "\FileManager"
        End Get
    End Property

    Public ReadOnly Property DefaultSettingFileName As String
        Get
            Return Me.ProjectDirectoryName & "\defaultSetting.json"
        End Get
    End Property



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

    Protected Function ProjectFile(ByVal f As String) As String
        If f.Contains(Me.ProjectDirectoryName) Then
            ProjectFile = f
        Else
            ProjectFile = Me.ProjectDirectoryName & "\" & f
        End If
    End Function

    Public Function _ProjectDirectory(ByVal d As String) As String
        _ProjectDirectory = ProjectFile(d)
    End Function

    Public ReadOnly Property LoadHandler As Func(Of String, T)
        Get
            Return AddressOf Me.ModelLoad
        End Get
    End Property

    Public ReadOnly Property SaveHandler As Action(Of String, T)
        Get
            Return AddressOf Me.ModelSave
        End Get
    End Property

    Public Overloads Function ModelLoad(Of T)(ByVal f As String) As T
        Dim txt As String
        Dim sr As System.IO.StreamReader

        ModelLoad = Nothing

        Try
            f = ProjectFile(f)

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


    Public Overloads Function ModelLoad(Of T As Class)(ByVal f As String,
                                                       ByVal key As String) As T
        Dim txt As String
        Dim sr As System.IO.StreamReader
        Dim obj As Object

        ModelLoad = Nothing

        Try
            f = ProjectFile(f)

            sr = New System.IO.StreamReader(
                f, System.Text.Encoding.GetEncoding(SHIFT_JIS))
            txt = sr.ReadToEnd()

            obj = JsonConvert.DeserializeObject(txt)

            ' 一部をロード
            ModelLoad = obj(key)
        Catch ex As Exception
            '
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
        End Try
    End Function


    Public Overloads Sub ModelSave(Of T)(ByVal f As String,
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


    Public Overloads Sub ModelSave(Of T, T2)(ByVal f As String,
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


    Sub New()
        Dim mkdir As ProjectProxy : mkdir = AddressOf Me.DirectoryEstablish
        Dim mkfil As ProjectProxy : mkfil = AddressOf Me.FileEstablish

        Try
            If Not Directory.Exists(Me.ProjectDirectoryName) Then
                Call mkdir(Me.ProjectDirectoryName)
                Call mkfil(Me._ProjectIniFileName)
                Call mkfil(Me.FileManagerFileName)
                Call mkfil(Me.DefaultSettingFileName)
            Else
                If Not File.Exists(Me._ProjectIniFileName) Then
                    Throw New Exception
                Else
                    If Not File.Exists(Me.FileManagerFileName) Then
                        Call mkfil(Me.FileManagerFileName)
                    End If
                    If Not File.Exists(Me.DefaultSettingFileName) Then
                        Call mkfil(Me.DefaultSettingFileName)
                    End If
                End If
            End If
        Catch ex As Exception
            '
        Finally
            '
        End Try
    End Sub
End Class
