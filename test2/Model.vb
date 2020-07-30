Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports Newtonsoft.Json
Imports System.IO

Delegate Sub mySub()
Delegate Function myBoolFunction() As Boolean
Delegate Function myStringFunction() As String
Delegate Function myModelLoad(Of T)(ByVal f As String) As T
Delegate Sub mySubWithT(Of T)(ByVal t As T)

Public MustInherit Class ModelBase
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub
End Class


Public Class TreeViewM
    'TreeView は考え方が複雑・・・
    '描画線用のデータ構造ではTreeViewのリスト型
    Private Property _RealName As String
    Public Property RealName As String
        Get
            Return _RealName
        End Get
        Set(value As String)
            _RealName = value
        End Set
    End Property

    Private _Child As List(Of TreeViewM)
    Public Property Child As List(Of TreeViewM)
        Get
            Return _Child
        End Get
        Set(value As List(Of TreeViewM))
            _Child = value
        End Set
    End Property
    '-----------------------------------------------------------------------------'

    'Private _SaveName As String
    'Public Property SaveName As String
    '    Get
    '        Return _SaveName
    '    End Get
    '    Set(value As String)
    '        _SaveName = value
    '    End Set
    'End Property

    'Private _FieldName As String
    'Public Property FieldName As String
    '    Get
    '        Return _FieldName
    '    End Get
    '    Set(value As String)
    '        _FieldName = value
    '    End Set
    'End Property


    'Private _SourceValue As String
    'Public Property SourceValue As String
    '    Get
    '        Return _SourceValue
    '    End Get
    '    Set(value As String)
    '        _SourceValue = value
    '    End Set
    'End Property

    'Private _DistinationValue As String
    'Public Property DistinationValue As String
    '    Get
    '        Return _DistinationValue
    '    End Get
    '    Set(value As String)
    '        _DistinationValue = value
    '    End Set
    'End Property

    'Private _ServerName As String
    'Public Property ServerName As String
    '    Get
    '        Return _ServerName
    '    End Get
    '    Set(value As String)
    '        _ServerName = value
    '    End Set
    'End Property

    'Private Property _DataBaseName As String
    'Public Property DataBaseName As String
    '    Get
    '        Return _DataBaseName
    '    End Get
    '    Set(value As String)
    '        _DataBaseName = value
    '    End Set
    'End Property

    'Private Property _DataTableName As String
    'Public Property DataTableName As String
    '    Get
    '        Return _DataTableName
    '    End Get
    '    Set(value As String)
    '        _DataTableName = value
    '    End Set
    'End Property
End Class


Public Class MainModel
    Public ServerName As String
    Public DataBaseName As String
    Public DataTableName As String
    Public FieldName As String
    Public SourceValue As String
    Public DistinationValue As String
    '-- Field Name ------------------------------------------------------'
    'Private _FieldName As String
    'Public Property FieldName As String
    '    Get
    '        Return _FieldName
    '    End Get
    '    Set(value As String)
    '        _FieldName = value
    '    End Set
    'End Property
    '--------------------------------------------------------------------'

    '-- Source Value ----------------------------------------------------'
    'Private _SourceValue As String
    'Public Property SourceValue As String
    '    Get
    '        Return _SourceValue
    '    End Get
    '    Set(value As String)
    '        _SourceValue = value
    '    End Set
    'End Property
    ''--------------------------------------------------------------------'

    ''-- Distination Value -----------------------------------------------'
    'Private _DistinationValue As String
    'Public Property DistinationValue As String
    '    Get
    '        Return _DistinationValue
    '    End Get
    '    Set(value As String)
    '        _DistinationValue = value
    '    End Set
    'End Property
    ''--------------------------------------------------------------------'

    ''-- Server Name -----------------------------------------------------'
    'Private _ServerName As String
    'Public Property ServerName As String
    '    Get
    '        Return _ServerName
    '    End Get
    '    Set(value As String)
    '        _ServerName = value
    '    End Set
    'End Property
    ''--------------------------------------------------------------------'

    ''-- DataBase Name ---------------------------------------------------'
    'Private Property _DataBaseName As String
    'Public Property DataBaseName As String
    '    Get
    '        Return _DataBaseName
    '    End Get
    '    Set(value As String)
    '        _DataBaseName = value
    '    End Set
    'End Property
    ''--------------------------------------------------------------------'


    ''-- DataTable Name --------------------------------------------------'
    'Private Property _DataTableName As String
    'Public Property DataTableName As String
    '    Get
    '        Return _DataTableName
    '    End Get
    '    Set(value As String)
    '        _DataTableName = value
    '    End Set
    'End Property
    ''--------------------------------------------------------------------'
End Class

Public Class MyProject
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

    Public ReadOnly Property ConfigFileName As String
        Get
            Return _LoadConfigFileName()
        End Get
    End Property

    Public Function LoadConfigFileName() As String
        Dim myload As myModelLoad(Of FileManagerFileM)

        myload = AddressOf ModelLoad(Of FileManagerFileM)
        LoadConfigFileName = (myload("FileManagerFile")).ConfigFileName
        'Dim txt As String
        'Dim sr As System.IO.StreamReader

        'Dim ms As mySub
        'Dim mbf As myBoolFunction

        '_LoadConfigFileName = vbNullString
        'Try
        '    ms = Sub()
        '             sr = New System.IO.StreamReader(
        '                Me._FileManagerFileName, System.Text.Encoding.GetEncoding(SHIFT_JIS))
        '             txt = sr.ReadToEnd()
        '             _LoadConfigFileName = JsonConvert.DeserializeObject(Of FileManagerFileM)(txt).ConfigFileName
        '         End Sub
        '    mbf = Function()
        '              Return _FileCheck(Me._FileManagerFileName)
        '          End Function

        '    Select Case Me._MyFileState(mbf)
        '        Case 0
        '            Call ms()
        '        Case 1
        '        Case 99
        '        Case 999
        '    End Select
        'Catch ex As Exception
        '    '
        'Finally
        '    If sr IsNot Nothing Then
        '        sr.Close()
        '        sr = Nothing
        '    End If
        'End Try
    End Function

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

    Private Function _FileManagerFileInitialize() As FileManagerFileM
        Dim fmfm As New FileManagerFileM
        fmfm.ConfigFileName = Me._MyDirectoryName & "\" & CONFIG_FILE_JSON
        _FileManagerFileInitialize = fmfm
    End Function

    Public Function ModelLoad(Of T)(ByVal f As String) As T
        Dim txt As String
        Dim sr As System.IO.StreamReader
        Dim mp As MyProject

        Dim ms As mySub

        ModelLoad = Nothing

        'いずれの例外も初期化
        Try
            mp = New MyProject
            ms = Sub()
                     sr = New System.IO.StreamReader(
                        f, System.Text.Encoding.GetEncoding(SHIFT_JIS))
                     txt = sr.ReadToEnd()
                     ModelLoad = JsonConvert.DeserializeObject(Of T)(txt)
                 End Sub

            Call ms()
        Catch ex As Exception
            '
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
            If mp IsNot Nothing Then
                mp = Nothing
            End If
        End Try
    End Function


    Public Sub ModelSave(Of T)(ByVal m As T)
        Dim txt As String
        Dim sw As System.IO.StreamWriter
        Dim ms As mySub
        Dim mp As MyProject

        Try
            mp = New MyProject
            ms = Sub()
                     txt = JsonConvert.SerializeObject(m)
                     sw = New System.IO.StreamWriter(
                        mp.ConfigFileName, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
                     )
                     sw.Write(txt)
                 End Sub
            Call ms()
        Catch ex As Exception
            '
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
            If mp IsNot Nothing Then
                mp = Nothing
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
                      Call FileWrite(Me.ConfigFileName,
                                     JsonConvert.SerializeObject(New Model))
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

Public Class FileManagerFileM : Inherits ModelBase
    Private _ConfigFileName As String
    Public Property ConfigFileName As String
        Get
            Return _ConfigFileName
        End Get
        Set(value As String)
            _ConfigFileName = value
            RaisePropertyChanged("ConfigFileName")
        End Set
    End Property
End Class

Public Class Model
    Private Const SHIFT_JIS = "SHIFT_JIS"




    '-- My Project ---------------------------------------------------------------'
    Private mp As MyProject
    '-----------------------------------------------------------------------------'

    '-- Tree View ----------------------------------------------------------------'
    Private _TreeViews As List(Of TreeViewM)
    Public Property TreeViews As List(Of TreeViewM)
        Get
            Return _TreeViews
        End Get
        Set(value As List(Of TreeViewM))
            _TreeViews = value
        End Set
    End Property

    Public Function TreeViewLoad() As List(Of TreeViewM)
        TreeViewLoad = Me.ModelLoad().TreeViews
    End Function
    '-----------------------------------------------------------------------------'

End Class
