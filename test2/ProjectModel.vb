Imports System.IO
Imports Newtonsoft.Json

Delegate Function ModelLoadProxy(Of T)(ByVal f As String) As T
Delegate Sub ModelSaveProxy(Of T)(ByVal f As String, ByVal m As T)
Public Delegate Function ProjectCheckProxy(ByVal f As String) As Boolean

Public Class ProjectModel
    Private Const SHIFT_JIS = "SHIFT_JIS"
    Private Const CONFIG_FILE_JSON = "ConfigFile.json"

    'Private ReadOnly Property _FileManagerFileName As String
    '    Get
    '        Return _MyDirectoryName & "\FileManager.json"
    '    End Get
    'End Property

    Private ReadOnly Property _ProjectDirectoryName As String
        Get
            Return System.Environment.GetEnvironmentVariable("USERPROFILE") & "\test"
        End Get
    End Property

    Private ReadOnly Property _ProjectIniFileName As String
        Get
            Return Me._ProjectDirectoryName & "\ProjectInit"
        End Get
    End Property






    '-- File Manager File ------------------------------------------------------------------'
    'Public Function LoadFileManagerFile() As FileManagerFileModel
    '    Dim myload As ModelLoadProxy(Of FileManagerFileModel)

    '    myload = AddressOf ModelLoad(Of FileManagerFileModel)
    '    LoadFileManagerFile = myload(Me._FileManagerFileName)
    'End Function



    'Private Function _FileManagerFileInitialize() As FileManagerFileModel
    '    Dim fmfm As New FileManagerFileModel
    '    fmfm.ConfigFileName = Me._MyDirectoryName & "\" & CONFIG_FILE_JSON
    '    _FileManagerFileInitialize = fmfm
    'End Function

    '---------------------------------------------------------------------------------------'



    '-- Tree Views -------------------------------------------------------------------------'
    '   SqlStatus内に統合のため廃止
    'Public Function LoadTreeViewModel() As List(Of TreeViewModel)
    '    Dim myload As ModelLoadProxy(Of ConfigFileModel)

    '    myload = AddressOf ModelLoad(Of ConfigFileModel)
    '    LoadTreeViewModel = myload(Me.LoadFileManagerFile.ConfigFileName).TreeViews
    'End Function
    '---------------------------------------------------------------------------------------'



    '-- SqlStatus --------------------------------------------------------------------------'
    'Public Function LoadSqlModel() As SqlModel
    '    Dim proxy As ModelLoadProxy(Of ConfigFileModel)

    '    proxy = AddressOf ModelLoad(Of ConfigFileModel)
    '    LoadSqlModel = proxy(Me.LoadFileManagerFile.ConfigFileName).SqlStatus
    'End Function
    '---------------------------------------------------------------------------------------'



    '最初以外にもその都度、ProjectDirectory, ProjectIniFileはチェックする
    Public Function ProjectCheck(ByVal proxy As ProjectCheckProxy, ByVal f As String) As Int32
        f = _ProjectFile(f)
        If Me.DirectoryCheck(Me._ProjectDirectoryName) Then
            If Me.FileCheck(Me._ProjectIniFileName) Then
                If (proxy(f)) Then
                    Return 0    '正常
                Else
                    Return 1    '対象ファイルが存在しない
                End If
            Else
                Return 99       '既に同名ディレクトリが存在
            End If
        Else
            Return 999          'ディレクトリが存在しない
        End If
    End Function

    'Public ReadOnly Property DirectoryCheck(d) As Boolean
    '    Get
    '        Try
    '            If System.IO.Directory.Exists(d) Then
    '                Return True
    '            Else
    '                Return False
    '            End If
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Get
    'End Property

    Public Function DirectoryCheck(ByVal d As String) As Boolean
        d = _ProjectDirectory(d)
        If Directory.Exists(d) Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Function FileCheck(ByVal f As String) As Boolean
        f = _ProjectFile(f)
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
        f = _ProjectFile(f)
        System.IO.File.Create(f)
    End Sub


    Public Sub FileWrite(ByVal f As String, ByVal txt As String)
        f = _ProjectFile(f)
        Dim sw As StreamWriter
        sw = New System.IO.StreamWriter(
            f, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
        )
        sw.Write(txt)
        sw.Close()
        sw.Dispose()
    End Sub

    Private Function _ProjectFile(ByVal f As String) As String
        If f.Contains(Me._ProjectDirectoryName) Then
            _ProjectFile = f
        Else
            _ProjectFile = Me._ProjectDirectoryName & "\" & f
        End If
    End Function

    Public Function _ProjectDirectory(ByVal d As String) As String
        _ProjectDirectory = _ProjectFile(d)
    End Function


    Public Function ModelLoad(Of T)(ByVal f As String) As T
        Dim txt As String
        Dim sr As System.IO.StreamReader

        ModelLoad = Nothing

        'いずれの例外も初期化
        Try
            f = _ProjectFile(f)
            sr = New System.IO.StreamReader(
                f, System.Text.Encoding.GetEncoding(SHIFT_JIS)
            )
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


    Public Sub ModelSave(Of T)(ByVal f As String, ByVal m As T)
        Dim txt As String
        Dim sw As System.IO.StreamWriter

        Try
            f = _ProjectFile(f)
            txt = JsonConvert.SerializeObject(m)
            sw = New System.IO.StreamWriter(
                f, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
            )
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


    Sub New()
        Dim proxy As ProjectCheckProxy
        proxy = AddressOf DirectoryCheck
        Try
            Select Case ProjectCheck(proxy, Me._ProjectDirectoryName)
                Case 0
                    Exit Select
                Case 99
                    Throw New Exception
                Case 999
                    Call DirectoryEstablish(Me._ProjectDirectoryName)
                    Call FileWrite(Me._ProjectIniFileName, vbNullString)
            End Select
        Catch ex As Exception
            '
        Finally
        End Try
    End Sub
End Class
