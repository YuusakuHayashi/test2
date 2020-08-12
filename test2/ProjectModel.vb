Imports System.IO
Imports Newtonsoft.Json

Delegate Function ModelLoadProxy(Of T)(ByVal f As String) As T

Public Class ProjectModel
    Private Const SHIFT_JIS = "SHIFT_JIS"
    Private Const CONFIG_FILE_JSON = "ConfigFile.json"

    Private ReadOnly Property _FileManagerFileName As String
        Get
            Return _MyDirectoryName & "\FileManager.json"
        End Get
    End Property

    Private ReadOnly Property _MyDirectoryName As String
        Get
            Return System.Environment.GetEnvironmentVariable("USERPROFILE") & "\test"
        End Get
    End Property

    Private ReadOnly Property _IniFileName As String
        Get
            Return _MyDirectoryName & "\IniFile"
        End Get
    End Property






    '-- File Manager File ------------------------------------------------------------------'
    Public Function LoadFileManagerFile() As FileManagerFileModel
        Dim myload As ModelLoadProxy(Of FileManagerFileModel)

        myload = AddressOf ModelLoad(Of FileManagerFileModel)
        LoadFileManagerFile = myload(Me._FileManagerFileName)
    End Function



    Private Function _FileManagerFileInitialize() As FileManagerFileModel
        Dim fmfm As New FileManagerFileModel
        fmfm.ConfigFileName = Me._MyDirectoryName & "\" & CONFIG_FILE_JSON
        _FileManagerFileInitialize = fmfm
    End Function

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
    Public Function LoadSqlModel() As SqlModel
        Dim proxy As ModelLoadProxy(Of ConfigFileModel)

        proxy = AddressOf ModelLoad(Of ConfigFileModel)
        LoadSqlModel = proxy(Me.LoadFileManagerFile.ConfigFileName).SqlStatus
    End Function
    '---------------------------------------------------------------------------------------'


    Private Function _MyProjectState() As Int32
        'ディレクトリおよび、iniFileのチェック
        Dim mbf As myBoolFunction
        mbf = Function()
                  Return True
              End Function
        _MyProjectState = _MyFileState(mbf)
    End Function

    Private Function _MyFileState(ByVal d As myBoolFunction) As Int32
        If _DirectoryCheck(Me._MyDirectoryName) Then
            If _FileCheck(Me._IniFileName) Then
                If (d()) Then
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

    Private ReadOnly Property _DirectoryCheck(d) As Boolean
        Get
            Try
                If System.IO.Directory.Exists(d) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Get
    End Property

    Private Function _FileCheck(ByVal f As String) As Boolean
        If File.Exists(f) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub DirectoryEstablish(ByVal d As String)
        System.IO.Directory.CreateDirectory(d)
    End Sub

    Private Sub FileEstablish(ByVal f As String)
        System.IO.File.Create(f)
    End Sub

    Private Sub FileWrite(ByVal f As String, ByVal txt As String)
        Dim sw As StreamWriter
        sw = New System.IO.StreamWriter(
            f, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
        )
        sw.Write(txt)
        sw.Close()
        sw.Dispose()
    End Sub


    Public Function ModelLoad(Of T)(ByVal f As String) As T
        Dim txt As String
        Dim sr As System.IO.StreamReader

        ModelLoad = Nothing

        'いずれの例外も初期化
        Try
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
        Dim ms As mySub
        Dim ms2 As mySub
        Dim ms3 As mySub
        Dim ms4 As mySub


        Try
            ms = Sub()
                     Call DirectoryEstablish(Me._MyDirectoryName)
                 End Sub
            ms2 = Sub()
                      Call FileWrite(Me._IniFileName, vbNullString)
                  End Sub
            ms3 = Sub()
                      Call FileWrite(Me._FileManagerFileName,
                                     JsonConvert.SerializeObject(_FileManagerFileInitialize()))
                  End Sub
            ms4 = Sub()
                      Call FileWrite(Me.LoadFileManagerFile.ConfigFileName,
                                     JsonConvert.SerializeObject(New ConfigFileModel))
                  End Sub
            Select Case (Me._MyProjectState())
                Case 0
                    Exit Select
                Case 99
                    Throw New Exception
                Case 999
                    Call ms()
                    Call ms2()
                    Call ms3()
                    Call ms4()
            End Select
        Catch ex As Exception
            '
        Finally
        End Try
    End Sub
End Class
