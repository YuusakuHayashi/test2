Imports System.Collections.ObjectModel
Imports Newtonsoft.Json

Delegate Sub mySub()

Public Class ServerModel
    Private _SaveName As String
    Public Property SaveName As String
        Get
            Return _SaveName
        End Get
        Set(value As String)
            _SaveName = value
        End Set
    End Property

    Private _FieldName As String
    Public Property FieldName As String
        Get
            Return _FieldName
        End Get
        Set(value As String)
            _FieldName = value
        End Set
    End Property


    Private _SourceValue As String
    Public Property SourceValue As String
        Get
            Return _SourceValue
        End Get
        Set(value As String)
            _SourceValue = value
        End Set
    End Property

    Private _DistinationValue As String
    Public Property DistinationValue As String
        Get
            Return _DistinationValue
        End Get
        Set(value As String)
            _DistinationValue = value
        End Set
    End Property

    Private _ServerName As String
    Public Property ServerName As String
        Get
            Return _ServerName
        End Get
        Set(value As String)
            _ServerName = value
        End Set
    End Property

    Private Property _DataBaseName As String
    Public Property DataBaseName As String
        Get
            Return _DataBaseName
        End Get
        Set(value As String)
            _DataBaseName = value
        End Set
    End Property

    Private Property _DataTableName As String
    Public Property DataTableName As String
        Get
            Return _DataTableName
        End Get
        Set(value As String)
            _DataTableName = value
        End Set
    End Property
    Private Property _RealName As String
    Public Property RealName As String
        Get
            Return _RealName
        End Get
        Set(value As String)
            _RealName = value
        End Set
    End Property

    Private _Child As List(Of ServerModel)
    Public Property Child As List(Of ServerModel)
        Get
            Return _Child
        End Get
        Set(value As List(Of ServerModel))
            _Child = value
        End Set
    End Property
End Class

Public Class Model
    Private Const SHIFT_JIS = "SHIFT_JIS"

    Private _Servers As List(Of ServerModel)
    Public Property Servers As List(Of ServerModel)
        Get
            Return _Servers
        End Get
        Set(value As List(Of ServerModel))
            _Servers = value
        End Set
    End Property

    Public Function ServersLoad() As List(Of ServerModel)
        ServersLoad = ModelLoad().Servers
    End Function

    Private ReadOnly Property ConfigFile As String
        Get
            Return MyDirectory & "\ConfigFile.json"
        End Get
    End Property

    Private ReadOnly Property IniFile As String
        Get
            Return MyDirectory & "\IniFile"
        End Get
    End Property

    Private ReadOnly Property MyDirectory As String
        Get
            Return System.Environment.GetEnvironmentVariable("USERPROFILE") & "\test"
        End Get
    End Property

    Private Sub MyDirectoryEstablish()
        System.IO.Directory.CreateDirectory(Me.MyDirectory)
    End Sub

    Private Sub IniFileEstablish()
        System.IO.File.Create(Me.IniFile)
    End Sub

    Private Sub ConfigFileEstablish()
        System.IO.File.Create(Me.ConfigFile)
    End Sub

    Private Function ModelLoad() As Model
        Dim txt As String
        Dim sr As System.IO.StreamReader
        Dim res As Int32

        ModelLoad = Nothing

        'いずれの例外も初期化
        Try
            res = _MyDirectoryState
            Select Case res
                Case 0
                    sr = New System.IO.StreamReader(
                        Me.ConfigFile, System.Text.Encoding.GetEncoding(SHIFT_JIS)
                    )
                    txt = sr.ReadToEnd()
                    ModelLoad = JsonConvert.DeserializeObject(Of Model)(txt)
                Case 1, 99, 999
                    Throw New Exception
                Case Else
                    Throw New Exception
            End Select

        Catch ex As Exception
            '
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr = Nothing
            End If
        End Try
    End Function

    Public Sub ModelSave()
        Dim txt As String
        Dim sw As System.IO.StreamWriter

        Dim ms As mySub
        ms = Sub()
                 txt = JsonConvert.SerializeObject(Me)
                 sw = New System.IO.StreamWriter(
                    Me.ConfigFile, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
                 )
                 sw.WriteLine(txt)
             End Sub
        Dim ms2 As mySub : ms2 = AddressOf MyDirectoryEstablish
        Dim ms3 As mySub : ms3 = AddressOf IniFileEstablish

        Dim res As Int32 : res = _MyDirectoryState
        Try
            Select Case res
                Case 0, 1
                    Call ms()
                Case 99
                    '
                Case 999
                    Call ms2()
                    Call ms3()
                    Call ms()
                Case Else
                    '
            End Select
        Catch ex As Exception
            '
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
        End Try
    End Sub

    Private ReadOnly Property _MyDirectoryState As Int32
        Get
            If _DirectoryCheck(Me.MyDirectory) Then
                If _FileCheck(Me.IniFile) Then
                    If _FileCheck(Me.ConfigFile) Then
                        Return 0    '正常
                    Else
                        Return 1    'Configファイルが存在せず、初回に発生しうる
                    End If
                Else
                    Return 99       '既に同名ファイルが存在
                End If
            Else
                Return 999          'ディレクトリが存在しない
            End If
        End Get
    End Property


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

    Private ReadOnly Property _FileCheck(ByVal f As String) As Boolean
        Get
            Try
                If System.IO.File.Exists(f) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Get
    End Property
End Class