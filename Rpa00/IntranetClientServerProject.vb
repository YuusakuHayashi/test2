Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Forms
Imports Rpa00

'Public Class RpaProject : Inherits RpaCui2.JsonHandler(Of RpaProject)
Public Class IntranetClientServerProject
    Inherits RpaProjectBase(Of IntranetClientServerProject)

    Private Const SHIFT_JIS As String = "Shift-JIS"
    Private Const UTF8 As String = "utf-8"
    Public Shared MYENCODING As String = UTF8

    Private Const ARCHITECTURE_NAME As String = "IntranetClientServer"

    <JsonIgnore>
    Public PrivateMode As Boolean

    Private ReadOnly Property HelloFileName As String
        Get
            Return $"{IntranetClientServerProject.ServerDirectory}\Hello.txt"
        End Get
    End Property

    Private Shared _RootDirectory As String
    Public Property RootDirectory As String
        Get
            Return IntranetClientServerProject._RootDirectory
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(IntranetClientServerProject._RootDirectory) Then
                IntranetClientServerProject._RootDirectory = value
            End If
        End Set
    End Property

    Private Shared _ServerDirectory As String
    Public Shared Property ServerDirectory As String
        Get
            Return IntranetClientServerProject._ServerDirectory
        End Get
        Set(value As String)
            If Not String.IsNullOrEmpty(value) Then
                IntranetClientServerProject._ServerDirectory = value
            End If
        End Set
    End Property

    ' Root Directory 関係
    '-----------------------------------------------------------------------------------'
    Public ReadOnly Property RootRobotDirectory As String
        Get
            Return $"{Me.RootDirectory}\{Me.RobotName}"
        End Get
    End Property

    Public ReadOnly Property IsRootRobotExists As Boolean
        Get
            If Directory.Exists(Me.RootRobotDirectory) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property RootRobotJsonFileName As String
        Get
            If Me.IsRootRobotExists Then
                Return $"{Me.RootRobotDirectory}\rpa_project.json"
            Else
                Return vbNullString
            End If
        End Get
    End Property

    Public ReadOnly Property RootRobotIniFileName As String
        Get
            Return $"{Me.RootRobotDirectory}\robot.ini"
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
            Return $"{Me.RootRobotDirectory}\ignore"
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
            Return $"{Me.MyRobotDirectory}\ignore"
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
            Return $"{CommonProject.SystemDirectory}\{Me.GetType.Name}"
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

    Public Overrides ReadOnly Property SystemJsonChangeFileName As String
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
            sr = New StreamReader(Me.HelloFileName, Encoding.GetEncoding(MYENCODING))
            txt = sr.ReadToEnd()
            Call _ConsoleWriteLine($"{txt}")
        End If

        If sr IsNot Nothing Then
            sr.Close()
            sr.Dispose()
        End If
    End Sub

    Private Sub _CheckProject()
        Call _ConsoleWriteLine($"プロジェクト '{Me.RobotAlias}' のチェック....")
        If Not Me.IsRootRobotExists Then
            Call _ConsoleWriteLine($"RootRobotDirectory '{Me.RootRobotDirectory}' がありません")
        End If
        If Not Me.IsMyRobotExists Then
            Call _ConsoleWriteLine($"MyRobotDirectory '{Me.MyRobotDirectory}' がありません")
        End If
        Call _ConsoleWriteLine("プロジェクトのチェック終了")

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

    'Public Overrides Sub CheckProject()
    '    Dim [changed] As Boolean = False
    '    Dim [old] As String = vbNullString
    '    If Not Directory.Exists(Me.RootDirectory) Then
    '        [old] = Me.RootDirectory
    '        Call _ConsoleWriteLine($"RootDirectory '{Me.RootDirectory}' がありません")
    '        Call _CreateProjectDirectory("RootDirectory")
    '        If [old] <> Me.RootDirectory Then
    '            [changed] = True
    '        End If
    '    End If
    '    'If Not Directory.Exists(Me.MyDirectory) Then
    '    '    [old] = Me.MyDirectory
    '    '    Call _ConsoleWriteLine($"MyDirectory '{Me.MyDirectory}' がありません")
    '    '    Call _CreateProjectDirectory("MyDirectory")
    '    '    If [old] <> Me.RootDirectory Then
    '    '        [changed] = True
    '    '    End If
    '    'End If
    '    If [changed] Then
    '        Call Save(Me.SystemJsonFileName, Me)
    '    End If
    'End Sub

    'Private Sub _CreateProjectDirectory(ByVal dtype As String)
    '    Dim yorn As String = vbNullString
    '    Dim yorn2 As String = vbNullString
    '    Dim fbd As FolderBrowserDialog
    '    Do
    '        yorn = vbNullString
    '        Console.WriteLine($"{dtype} の設定を行いますか (y/n)")
    '        Console.Write($">>> ")
    '        yorn = Console.ReadLine()
    '    Loop Until yorn = "y" Or yorn = "n"
    '    If yorn = "y" Then
    '        Do
    '            yorn2 = vbNullString
    '            fbd = New FolderBrowserDialog With {
    '                .Description = $"Select {dtype}",
    '                .RootFolder = Environment.SpecialFolder.Desktop,
    '                .SelectedPath = Environment.SpecialFolder.Desktop,
    '                .ShowNewFolderButton = True
    '            }
    '            If fbd.ShowDialog() = DialogResult.OK Then
    '                Console.WriteLine($"よろしいですか？ '{fbd.SelectedPath}' (y/n)")
    '                Console.Write($">>> ")
    '                yorn2 = Console.ReadLine()
    '            Else
    '                yorn2 = "x"
    '            End If
    '        Loop Until yorn2 = "y" Or yorn2 = "x"
    '        If yorn2 = "y" Then
    '            If dtype = "RootDirectory" Then
    '                Me.RootDirectory = fbd.SelectedPath
    '            End If
    '            If dtype = "MyDirectory" Then
    '                Me.MyDirectory = fbd.SelectedPath
    '            End If
    '            Console.WriteLine($"{dtype} が初期設定されました")
    '        End If
    '        If yorn2 = "x" Then
    '            Console.WriteLine($"{dtype} の初期設定は行いませんでした")
    '        End If
    '    End If
    'End Sub

    Private Sub _ConsoleWriteLine(ByVal [line] As String)
        If Not Me.PrivateMode Then
            Console.WriteLine([line])
        End If
    End Sub

    Public Sub RunShell(ByVal exe As String, ByVal arg As String)
        Dim proc = New System.Diagnostics.Process()
        'proc.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
        proc.StartInfo.FileName = exe
        proc.StartInfo.UseShellExecute = False
        proc.StartInfo.RedirectStandardOutput = True
        proc.StartInfo.RedirectStandardInput = False
        proc.StartInfo.CreateNoWindow = True
        proc.StartInfo.Arguments = arg
        proc.Start()
        proc.WaitForExit()
        proc.Close()
    End Sub

    Private Function _GetIgnoreList(ByVal f As String) As List(Of String)
        Dim txt As String
        Dim sr As StreamReader
        Try
            sr = New System.IO.StreamReader(
                f, System.Text.Encoding.GetEncoding(MYENCODING))
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

    Protected Overrides Sub CreateChangedFile()
        If Me.FirstLoad Then
            Call RpaModule.CreateChangedFile(Me.SystemJsonChangeFileName)
        End If
    End Sub
End Class
