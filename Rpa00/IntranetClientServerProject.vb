Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Forms

'Public Class RpaProject : Inherits RpaCui2.JsonHandler(Of RpaProject)
Public Class IntranetClientServerProject : Inherits RpaProjectBase(Of IntranetClientServerProject)

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
    Public ReadOnly Property RootProjectDirectory As String
        Get
            Return $"{Me.RootDirectory}\{Me.ProjectName}"
        End Get
    End Property

    Public ReadOnly Property IsRootProjectExists As Boolean
        Get
            If Directory.Exists(Me.RootProjectDirectory) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property RootProjectJsonFileName As String
        Get
            If Me.IsRootProjectExists Then
                Return $"{Me.RootProjectDirectory}\rpa_project.json"
            Else
                Return vbNullString
            End If
        End Get
    End Property

    Private Property _RootProjectObject As Object
    <JsonIgnore>
    Public Property RootProjectObject As Object
        Get
            If Me._RootProjectObject Is Nothing Then
                Dim obj = RpaCodes.RpaObject(Me)
                Dim obj2 = Nothing
                If obj Is Nothing Then
                    Me._RootProjectObject = Nothing
                Else
                    If File.Exists(Me.RootProjectJsonFileName) Then
                        obj2 = obj.Load(Me.RootProjectJsonFileName)
                    End If
                    If obj2 Is Nothing Then
                        Me._RootProjectObject = obj
                    Else
                        Me._RootProjectObject = obj2
                    End If
                End If
            End If
            Return Me._RootProjectObject
        End Get
        Set(value As Object)
            Me._RootProjectObject = value
        End Set
    End Property

    Public ReadOnly Property RootProjectIgnoreFileName As String
        Get
            Return $"{Me.RootProjectDirectory}\ignore"
        End Get
    End Property

    ' セーブ？ロード？するたびに、リストが重複するためＩＧＮＯＲＥするようにした。
    ' 原因が分かったら、ＩＧＮＯＲＥは解除する？しなくてもいいか・・・
    Private _RootProjectIgnoreList As List(Of String)
    <JsonIgnore>
    Public Property RootProjectIgnoreList As List(Of String)
        Get
            If Me._RootProjectIgnoreList Is Nothing Then
                Me._RootProjectIgnoreList = New List(Of String)
                If File.Exists(Me.RootProjectIgnoreFileName) Then
                    Me._RootProjectIgnoreList = _GetIgnoreList(Me.RootProjectIgnoreFileName)
                End If
            End If
            Return Me._RootProjectIgnoreList
        End Get
        Set(value As List(Of String))
            Me._RootProjectIgnoreList = value
        End Set
    End Property

    'Public ReadOnly Property RootProjectUpdateDirectory As String
    '    Get
    '        Return Me.RootProjectDirectory & "\updates"
    '    End Get
    'End Property

    'Private _RootProjectUpdatePackages As List(Of RpaPackage)
    'Public ReadOnly Property RootProjectUpdatePackages As List(Of RpaPackage)
    '    Get
    '        Dim pack As RpaPackage
    '        Dim jh As New RpaCui2.JsonHandler(Of Object)
    '        If Me._RootProjectUpdatePackages Is Nothing Then
    '            Me._RootProjectUpdatePackages = New List(Of RpaPackage)
    '            If Directory.Exists(Me.RootProjectUpdateDirectory) Then
    '                For Each d In Directory.GetDirectories(Me.RootProjectUpdateDirectory)
    '                    If Directory.GetFiles(d).Contains(d & "\RpaPackage.json") Then
    '                        pack = jh.Load(Of RpaPackage)(d & "\RpaPackage.json")
    '                        If pack Is Nothing Then
    '                            Me._RootProjectUpdatePackages.Add(pack)
    '                        End If
    '                    End If
    '                Next
    '            End If
    '        End If
    '        Return Me._RootProjectUpdatePackages
    '    End Get
    'End Property
    '-----------------------------------------------------------------------------------'

    ' MyProject Directory 関係
    '-----------------------------------------------------------------------------------'
    'Public ReadOnly Property IsMyProjectExists As Boolean
    '    Get
    '        If Directory.Exists(Me.MyProjectDirectory) Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    End Get
    'End Property

    'Public ReadOnly Property MyProjectDirectory As String
    '    Get
    '        Return $"{Me.MyDirectory}\{Me.ProjectAlias}"
    '    End Get
    'End Property

    'Private _MyProjectJsonFileName As String
    'Public ReadOnly Property MyProjectJsonFileName As String
    '    Get
    '        If String.IsNullOrEmpty(Me._MyProjectJsonFileName) Then
    '            Me._MyProjectJsonFileName = $"{Me.MyProjectDirectory}\rpa_project.json"
    '        End If
    '        Return Me._MyProjectJsonFileName
    '    End Get
    'End Property

    'Private Property _MyProjectObject As Object
    '<JsonIgnore>
    'Public Property MyProjectObject As Object
    '    Get
    '        If Me._MyProjectObject Is Nothing Then
    '            Dim obj = RpaCodes.RpaObject(Me)
    '            Dim obj2 = Nothing
    '            If obj Is Nothing Then
    '                Me._MyProjectObject = Nothing
    '            Else
    '                If File.Exists(Me.MyProjectJsonFileName) Then
    '                    obj2 = obj.Load(Me.MyProjectJsonFileName)
    '                End If
    '                If obj2 Is Nothing Then
    '                    Me._MyProjectObject = obj
    '                Else
    '                    Me._MyProjectObject = obj2
    '                End If
    '            End If
    '        End If
    '        Return Me._MyProjectObject
    '    End Get
    '    Set(value As Object)
    '        Me._MyProjectObject = value
    '    End Set
    'End Property

    Public ReadOnly Property MyProjectIgnoreFileName As String
        Get
            Return $"{Me.MyProjectDirectory}\ignore"
        End Get
    End Property

    Private _MyProjectIgnoreList As List(Of String)
    <JsonIgnore>
    Public Property MyProjectIgnoreList As List(Of String)
        Get
            If Me._MyProjectIgnoreList Is Nothing Then
                Me._MyProjectIgnoreList = New List(Of String)
                If File.Exists(Me.MyProjectIgnoreFileName) Then
                    Me._MyProjectIgnoreList = _GetIgnoreList(Me.MyProjectIgnoreFileName)
                End If
            End If
            Return Me._MyProjectIgnoreList
        End Get
        Set(value As List(Of String))
            Me._MyProjectIgnoreList = value
        End Set
    End Property

    Public Overrides ReadOnly Property SystemSolutionDirectory As String
        Get
            Return $"{Me.SystemArchitecutureDirectory}\{Me.SolutionName}"
        End Get
    End Property

    Public Overrides ReadOnly Property SystemJsonFileName As String
        Get
            Return $"{Me.SystemSolutionDirectory}\rpa_project.json"
        End Get
    End Property

    Public Overrides ReadOnly Property SystemArchitecutureDirectory As String
        Get
            Return $"{CommonProject.SystemDirectory}\{ARCHITECTURE_NAME}"
        End Get
    End Property

    'Public ReadOnly Property MyProjectUpdatedDirectory As String
    '    Get
    '        Return Me.MyProjectDirectory & "\updated"
    '    End Get
    'End Property

    'Private _MyProjectUpdatedPackages As List(Of RpaPackage)
    'Public ReadOnly Property MyProjectUpdatedPackages As List(Of RpaPackage)
    '    Get
    '        Dim pack As RpaPackage
    '        Dim jh As New RpaCui2.JsonHandler(Of Object)
    '        If Me._MyProjectUpdatedPackages Is Nothing Then
    '            Me._MyProjectUpdatedPackages = New List(Of RpaPackage)
    '            If Directory.Exists(Me.MyProjectUpdatedDirectory) Then
    '                For Each f In Directory.GetFiles(Me.MyProjectUpdatedDirectory)
    '                    If Path.GetFileName(f) = "RpaPackage.json" Then
    '                        pack = jh.Load(Of RpaPackage)(f)
    '                        If pack Is Nothing Then
    '                            Me._MyProjectUpdatedPackages.Add(pack)
    '                        End If
    '                    End If
    '                Next
    '            End If
    '        End If
    '        Return Me._MyProjectUpdatedPackages
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
        Call _ConsoleWriteLine($"プロジェクト '{Me.ProjectAlias}' のチェック....")
        If Not Me.IsRootProjectExists Then
            Call _ConsoleWriteLine($"RootProjectDirectory '{Me.RootProjectDirectory}' がありません")
        End If
        If Not Me.IsMyProjectExists Then
            Call _ConsoleWriteLine($"MyProjectDirectory '{Me.MyProjectDirectory}' がありません")
        End If
        Call _ConsoleWriteLine("プロジェクトのチェック終了")

        'Dim installed = False
        'Call _ConsoleWriteLine("アップデートパッケージを検索しています...")
        'For Each rp In Me.RootProjectUpdatePackages
        '    installed = False
        '    For Each mp In Me.MyProjectUpdatedPackages
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

    Public Overrides Sub CheckProject()
        Dim [changed] As Boolean = False
        Dim [old] As String = vbNullString
        If Not Directory.Exists(Me.RootDirectory) Then
            [old] = Me.RootDirectory
            Call _ConsoleWriteLine($"RootDirectory '{Me.RootDirectory}' がありません")
            Call _CreateProjectDirectory("RootDirectory")
            If [old] <> Me.RootDirectory Then
                [changed] = True
            End If
        End If
        'If Not Directory.Exists(Me.MyDirectory) Then
        '    [old] = Me.MyDirectory
        '    Call _ConsoleWriteLine($"MyDirectory '{Me.MyDirectory}' がありません")
        '    Call _CreateProjectDirectory("MyDirectory")
        '    If [old] <> Me.RootDirectory Then
        '        [changed] = True
        '    End If
        'End If
        If [changed] Then
            Call Save(Me.SystemJsonFileName, Me)
        End If
    End Sub

    Private Sub _CreateProjectDirectory(ByVal dtype As String)
        Dim yorn As String = vbNullString
        Dim yorn2 As String = vbNullString
        Dim fbd As FolderBrowserDialog
        Do
            yorn = vbNullString
            Console.WriteLine($"{dtype} の設定を行いますか (y/n)")
            Console.Write($">>> ")
            yorn = Console.ReadLine()
        Loop Until yorn = "y" Or yorn = "n"
        If yorn = "y" Then
            Do
                yorn2 = vbNullString
                fbd = New FolderBrowserDialog With {
                    .Description = $"Select {dtype}",
                    .RootFolder = Environment.SpecialFolder.Desktop,
                    .SelectedPath = Environment.SpecialFolder.Desktop,
                    .ShowNewFolderButton = True
                }
                If fbd.ShowDialog() = DialogResult.OK Then
                    Console.WriteLine($"よろしいですか？ '{fbd.SelectedPath}' (y/n)")
                    Console.Write($">>> ")
                    yorn2 = Console.ReadLine()
                Else
                    yorn2 = "x"
                End If
            Loop Until yorn2 = "y" Or yorn2 = "x"
            If yorn2 = "y" Then
                If dtype = "RootDirectory" Then
                    Me.RootDirectory = fbd.SelectedPath
                End If
                If dtype = "MyDirectory" Then
                    Me.MyDirectory = fbd.SelectedPath
                End If
                Console.WriteLine($"{dtype} が初期設定されました")
            End If
            If yorn2 = "x" Then
                Console.WriteLine($"{dtype} の初期設定は行いませんでした")
            End If
        End If
    End Sub

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
            Call _ConsoleWriteLine(ex.Message)
            _GetIgnoreList = Nothing
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
        End Try
    End Function

    Public Sub New()
    End Sub

End Class
