Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Forms

'Public Class RpaProject : Inherits RpaCui.JsonHandler(Of RpaProject)
Public Class IntranetClientServerProject
    Inherits RpaProjectBase(Of IntranetClientServerProject)

    Private Const SHIFT_JIS As String = "Shift-JIS"
    Private Const UTF8 As String = "utf-8"

    Private Const ARCHITECTURE_NAME As String = "IntranetClientServer"

    <JsonIgnore>
    Public PrivateMode As Boolean

    Private ReadOnly Property HelloFileName As String
        Get
            Return $"{Me.ServerDirectory}\Hello.txt"
        End Get
    End Property

    Private _RootDirectory As String
    Public Property RootDirectory As String
        Get
            Return Me._RootDirectory
        End Get
        Set(value As String)
            Me._RootDirectory = value
            RaisePropertyChanged("RootDirectory")
        End Set
    End Property

    Private _ServerDirectory As String
    Public Property ServerDirectory As String
        Get
            Return Me._ServerDirectory
        End Get
        Set(value As String)
            Me._ServerDirectory = value
            RaisePropertyChanged("ServerDirectory")
        End Set
    End Property

    ' Root Directory 関係
    '-----------------------------------------------------------------------------------'
    <JsonIgnore>
    Public ReadOnly Property RootRobotDirectory As String
        Get
            Dim [dir] As String = vbNullString
            If Not String.IsNullOrEmpty(Me.RootDirectory) Then
                If Directory.Exists(Me.RootDirectory) Then
                    If Not String.IsNullOrEmpty(Me.RobotName) Then
                        [dir] = $"{Me.RootDirectory}\{Me.RobotName}"
                    End If
                End If
            End If
            Return [dir]
        End Get
    End Property

    <JsonIgnore>
    Public ReadOnly Property RootRobotJsonFileName As String
        Get
            Dim [fil] As String = vbNullString
            If Directory.Exists(Me.RootRobotDirectory) Then
                [fil] = $"{Me.RootRobotDirectory}\rpa_project.json"
            End If
            Return [fil]
        End Get
    End Property

    ' このファイルがRootRobotになければ、AttachRobotすることができない
    <JsonIgnore>
    Public ReadOnly Property RootRobotIniFileName As String
        Get
            Dim [fil] As String = vbNullString
            If Directory.Exists(Me.RootRobotDirectory) Then
                [fil] = $"{Me.RootRobotDirectory}\robot.ini"
            End If
            Return [fil]
        End Get
    End Property

    <JsonIgnore>
    Public ReadOnly Property RootRobotRepositoryDirectory As String
        Get
            Dim [dir] As String = vbNullString
            If Directory.Exists(Me.RootRobotDirectory) Then
                [dir] = $"{Me.RootRobotDirectory}\repo"
                If Not Directory.Exists([dir]) Then
                    Directory.CreateDirectory([dir])
                End If
            End If
            Return [dir]
        End Get
    End Property

    <JsonIgnore>
    Public ReadOnly Property RootRobotDllRepositoryDirectory As String
        Get
            Dim [dir] As String = vbNullString
            If Directory.Exists(Me.RootRobotRepositoryDirectory) Then
                [dir] = $"{Me.RootRobotRepositoryDirectory}\dll"
                If Not Directory.Exists([dir]) Then
                    Directory.CreateDirectory([dir])
                End If
            End If
            Return [dir]
        End Get
    End Property

    Private Property _RootRobotObject As Object
    <JsonIgnore>
    Public Property RootRobotObject As Object
        Get
            If Me._RootRobotObject Is Nothing Then
                Dim obj = RpaCodes.RpaObject(Me)
                Dim obj2 = Nothing
                If obj Is Nothing Then
                    Me._RootRobotObject = Nothing
                Else
                    If File.Exists(Me.RootRobotJsonFileName) Then
                        obj2 = obj.Load(Me.RootRobotJsonFileName)
                    End If
                    If obj2 Is Nothing Then
                        Me._RootRobotObject = obj
                    Else
                        Me._RootRobotObject = obj2
                    End If
                End If
            End If
            Return Me._RootRobotObject
        End Get
        Set(value As Object)
            Me._RootRobotObject = value
        End Set
    End Property

    Public ReadOnly Property RootRobotIgnoreFileName As String
        Get
            Dim [fil] As String = vbNullString
            If Directory.Exists(Me.RootRobotDirectory) Then
                [fil] = $"{Me.RootRobotDirectory}\ignore"
            End If
            Return [fil]
        End Get
    End Property

    ' セーブ？ロード？するたびに、リストが重複するためＩＧＮＯＲＥするようにした。
    ' 原因が分かったら、ＩＧＮＯＲＥは解除する？しなくてもいいか・・・
    Private _RootRobotIgnoreList As List(Of String)
    <JsonIgnore>
    Public Property RootRobotIgnoreList As List(Of String)
        Get
            If Me._RootRobotIgnoreList Is Nothing Then
                If File.Exists(Me.RootRobotIgnoreFileName) Then
                    Me._RootRobotIgnoreList = _GetIgnoreList(Me.RootRobotIgnoreFileName)
                Else
                    Me._RootRobotIgnoreList = New List(Of String)
                End If
            End If
            Return Me._RootRobotIgnoreList
        End Get
        Set(value As List(Of String))
            Me._RootRobotIgnoreList = value
        End Set
    End Property

    'Public ReadOnly Property RootRobotUpdateDirectory As String
    '    Get
    '        Return Me.RootRobotDirectory & "\updates"
    '    End Get
    'End Property

    'Private _RootRobotUpdatePackages As List(Of RpaPackage)
    'Public ReadOnly Property RootRobotUpdatePackages As List(Of RpaPackage)
    '    Get
    '        Dim pack As RpaPackage
    '        Dim jh As New RpaCui2.JsonHandler(Of Object)
    '        If Me._RootRobotUpdatePackages Is Nothing Then
    '            Me._RootRobotUpdatePackages = New List(Of RpaPackage)
    '            If Directory.Exists(Me.RootRobotUpdateDirectory) Then
    '                For Each d In Directory.GetDirectories(Me.RootRobotUpdateDirectory)
    '                    If Directory.GetFiles(d).Contains(d & "\RpaPackage.json") Then
    '                        pack = jh.Load(Of RpaPackage)(d & "\RpaPackage.json")
    '                        If pack Is Nothing Then
    '                            Me._RootRobotUpdatePackages.Add(pack)
    '                        End If
    '                    End If
    '                Next
    '            End If
    '        End If
    '        Return Me._RootRobotUpdatePackages
    '    End Get
    'End Property
    '-----------------------------------------------------------------------------------'

    ' MyRobot Directory 関係
    '-----------------------------------------------------------------------------------'
    'Public ReadOnly Property IsMyRobotExists As Boolean
    '    Get
    '        If Directory.Exists(Me.MyRobotDirectory) Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    End Get
    'End Property

    'Public ReadOnly Property MyRobotDirectory As String
    '    Get
    '        Return $"{Me.MyDirectory}\{Me.RobotAlias}"
    '    End Get
    'End Property

    'Private _MyRobotJsonFileName As String
    'Public ReadOnly Property MyRobotJsonFileName As String
    '    Get
    '        If String.IsNullOrEmpty(Me._MyRobotJsonFileName) Then
    '            Me._MyRobotJsonFileName = $"{Me.MyRobotDirectory}\rpa_project.json"
    '        End If
    '        Return Me._MyRobotJsonFileName
    '    End Get
    'End Property

    'Private Property _MyRobotObject As Object
    '<JsonIgnore>
    'Public Property MyRobotObject As Object
    '    Get
    '        If Me._MyRobotObject Is Nothing Then
    '            Dim obj = RpaCodes.RpaObject(Me)
    '            Dim obj2 = Nothing
    '            If obj Is Nothing Then
    '                Me._MyRobotObject = Nothing
    '            Else
    '                If File.Exists(Me.MyRobotJsonFileName) Then
    '                    obj2 = obj.Load(Me.MyRobotJsonFileName)
    '                End If
    '                If obj2 Is Nothing Then
    '                    Me._MyRobotObject = obj
    '                Else
    '                    Me._MyRobotObject = obj2
    '                End If
    '            End If
    '        End If
    '        Return Me._MyRobotObject
    '    End Get
    '    Set(value As Object)
    '        Me._MyRobotObject = value
    '    End Set
    'End Property

    Public ReadOnly Property MyRobotIgnoreFileName As String
        Get
            Dim [fil] As String = vbNullString
            If Directory.Exists(Me.MyRobotDirectory) Then
                [fil] = $"{Me.MyRobotDirectory}\ignore"
            End If
            Return [fil]
        End Get
    End Property

    Private _MyRobotIgnoreList As List(Of String)
    <JsonIgnore>
    Public Property MyRobotIgnoreList As List(Of String)
        Get
            If Me._MyRobotIgnoreList Is Nothing Then
                If File.Exists(Me.MyRobotIgnoreFileName) Then
                    Me._MyRobotIgnoreList = _GetIgnoreList(Me.MyRobotIgnoreFileName)
                Else
                    Me._MyRobotIgnoreList = New List(Of String)
                End If
            End If
            Return Me._MyRobotIgnoreList
        End Get
        Set(value As List(Of String))
            Me._MyRobotIgnoreList = value
        End Set
    End Property

    Public Overrides ReadOnly Property SystemProjectDirectory As String
        Get
            Return $"{Me.SystemArchDirectory}\{Me.ProjectName}"
        End Get
    End Property

    Public Overrides ReadOnly Property SystemJsonFileName As String
        Get
            Return $"{Me.SystemProjectDirectory}\rpa_project.json"
        End Get
    End Property

    Public Overrides ReadOnly Property SystemArchDirectory As String
        Get
            Return $"{RpaCui.SystemDirectory}\{Me.GetType.Name}"
        End Get
    End Property

    Public Overrides ReadOnly Property SystemArchType As Integer
        Get
            Return 1
        End Get
    End Property

    Public Overrides ReadOnly Property SystemArchTypeName As String
        Get
            Return Me.GetType.Name
        End Get
    End Property

    Public Overrides ReadOnly Property SystemJsonChangedFileName As String
        Get
            Return $"{Me.SystemProjectDirectory}\rpa_project.json.chg"
        End Get
    End Property


    'Public ReadOnly Property MyRobotUpdatedDirectory As String
    '    Get
    '        Return Me.MyRobotDirectory & "\updated"
    '    End Get
    'End Property

    'Private _MyRobotUpdatedPackages As List(Of RpaPackage)
    'Public ReadOnly Property MyRobotUpdatedPackages As List(Of RpaPackage)
    '    Get
    '        Dim pack As RpaPackage
    '        Dim jh As New RpaCui2.JsonHandler(Of Object)
    '        If Me._MyRobotUpdatedPackages Is Nothing Then
    '            Me._MyRobotUpdatedPackages = New List(Of RpaPackage)
    '            If Directory.Exists(Me.MyRobotUpdatedDirectory) Then
    '                For Each f In Directory.GetFiles(Me.MyRobotUpdatedDirectory)
    '                    If Path.GetFileName(f) = "RpaPackage.json" Then
    '                        pack = jh.Load(Of RpaPackage)(f)
    '                        If pack Is Nothing Then
    '                            Me._MyRobotUpdatedPackages.Add(pack)
    '                        End If
    '                    End If
    '                Next
    '            End If
    '        End If
    '        Return Me._MyRobotUpdatedPackages
    '    End Get
    'End Property
    '-----------------------------------------------------------------------------------'

    'Private _PrinterName As String
    'Public Property PrinterName As String
    '    Get
    '        Return Me._PrinterName
    '    End Get
    '    Set(value As String)
    '        Me._PrinterName = value
    '    End Set
    'End Property

    'Private _UsePrinterName As String
    '<JsonIgnore>
    'Public Property UsePrinterName As String
    '    Get
    '        Return Me._UsePrinterName
    '    End Get
    '    Set(value As String)
    '        Me._UsePrinterName = value
    '    End Set
    'End Property

    Public Sub HelloProject()
        Dim txt As String
        Dim sr As StreamReader

        If File.Exists(Me.HelloFileName) Then
            sr = New StreamReader(Me.HelloFileName, Encoding.GetEncoding(SHIFT_JIS))
            txt = sr.ReadToEnd()
            Call _ConsoleWriteLine($"{txt}")
        End If

        If sr IsNot Nothing Then
            sr.Close()
            sr.Dispose()
        End If
    End Sub

    Private Sub _CheckProject()
        'Call _ConsoleWriteLine($"プロジェクト '{Me.RobotAlias}' のチェック....")
        'If Not Me.IsRootRobotExists Then
        '    Call _ConsoleWriteLine($"RootRobotDirectory '{Me.RootRobotDirectory}' がありません")
        'End If
        'If Not Me.IsMyRobotExists Then
        '    Call _ConsoleWriteLine($"MyRobotDirectory '{Me.MyRobotDirectory}' がありません")
        'End If
        'Call _ConsoleWriteLine("プロジェクトのチェック終了")

        'Dim installed = False
        'Call _ConsoleWriteLine("アップデートパッケージを検索しています...")
        'For Each rp In Me.RootRobotUpdatePackages
        '    installed = False
        '    For Each mp In Me.MyRobotUpdatedPackages
        '        If rp.Name = mp.Name Then
        '            installed = True
        '            Exit For
        '        End If
        '    Next
        '    If Not installed Then
        '        Call _ConsoleWriteLine("パッケージ : " & rp.Name & " をインストールしていません")
        '    End If
        'Next
    End Sub

    Private Sub _ConsoleWriteLine(ByVal [line] As String)
        If Not Me.PrivateMode Then
            Console.WriteLine([line])
        End If
    End Sub

    'Public Sub RunShell(ByVal exe As String, ByVal arg As String)
    '    Dim proc = New System.Diagnostics.Process()
    '    'proc.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
    '    proc.StartInfo.FileName = exe
    '    proc.StartInfo.UseShellExecute = False
    '    proc.StartInfo.RedirectStandardOutput = True
    '    proc.StartInfo.RedirectStandardInput = False
    '    proc.StartInfo.CreateNoWindow = True
    '    proc.StartInfo.Arguments = arg
    '    proc.Start()
    '    proc.WaitForExit()
    '    proc.Close()
    'End Sub

    Private Function _GetIgnoreList(ByVal f As String) As List(Of String)
        Dim txt As String
        Dim sr As StreamReader
        Try
            sr = New System.IO.StreamReader(
                f, System.Text.Encoding.GetEncoding(SHIFT_JIS))
            txt = sr.ReadToEnd()

            _GetIgnoreList = JsonConvert.DeserializeObject(Of List(Of String))(txt)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            _GetIgnoreList = New List(Of String)
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
        End Try
    End Function


    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        If String.IsNullOrEmpty(Me.RootDirectory) Then
            Console.WriteLine($"'RootDirectory' がセットされていません")
            Return False
        End If
        If Not Directory.Exists(Me.RootDirectory) Then
            Console.WriteLine($"ディレクトリ '{Me.RootDirectory}' は存在しません")
            Return False
        End If
        If String.IsNullOrEmpty(Me.MyDirectory) Then
            Console.WriteLine($"'MyDirectory' がセットされていません")
            Return False
        End If
        If Not Directory.Exists(Me.MyDirectory) Then
            Console.WriteLine($"ディレクトリ '{Me.MyDirectory}' は存在しません")
            Return False
        End If
        If String.IsNullOrEmpty(Me.RobotName) Then
            Console.WriteLine($"'RobotName' がセットされていません")
            Return False
        End If
        If String.IsNullOrEmpty(Me.RootRobotDirectory) Then
            Console.WriteLine($"'RootRobotDirectory' がセットされていません")
            Return False
        End If
        If Not Directory.Exists(Me.RootRobotDirectory) Then
            Console.WriteLine($"ディレクトリ '{Me.RootRobotDirectory}' は存在しません")
            Return False
        End If
        If String.IsNullOrEmpty(Me.MyRobotDirectory) Then
            Console.WriteLine($"'MyRobotDirectory' がセットされていません")
            Return False
        End If
        If Not Directory.Exists(Me.MyRobotDirectory) Then
            Console.WriteLine($"ディレクトリ '{Me.MyRobotDirectory}' は存在しません")
            Return False
        End If
        Return True
    End Function


    Sub New()
        Me.CanExecuteHandler = AddressOf Check
        AddHandler Me.PropertyChanged, AddressOf CreateChangedFile
    End Sub


    Protected Overrides Sub Finalize()
        RemoveHandler Me.PropertyChanged, AddressOf CreateChangedFile
        MyBase.Finalize()
    End Sub
End Class
