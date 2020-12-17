Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class RpaProject : Inherits RpaCui2.JsonHandler(Of RpaProject)
    Private Const SHIFT_JIS As String = "Shift-JIS"
    Private Const UTF8 As String = "utf-8"
    Public Shared MYENCODING As String = UTF8

    Private Shared _RootDirectory As String
    Public Property RootDirectory As String
        Get
            Return RpaProject._RootDirectory
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(RpaProject._RootDirectory) Then
                RpaProject._RootDirectory = value
                If Not Directory.Exists(RpaProject._RootDirectory) Then
                    Console.WriteLine("RootDirectory が存在しません")
                    Console.WriteLine("ファイル: " & Me.SystemJsonFileName & "の'RootDirectory' に任意のパスを書いてください")
                    Console.WriteLine("ファイルを保存した後、アプリケーションを再起動してください")
                End If
            End If
        End Set
    End Property

    Private ReadOnly Property _RootSystemDirectory As String
        Get
            Return $"{Me.RootDirectory}\system"
        End Get
    End Property

    Public Shared ReadOnly Property SYSTEM_DIRECTORY As String
        Get
            Return System.Environment.GetEnvironmentVariable("USERPROFILE") & "\rpa_project"
        End Get
    End Property

    Public Shared ReadOnly Property SYSTEM_SCRIPT_DIRECTORY As String
        Get
            Dim d = $"{RpaProject.SYSTEM_DIRECTORY}\script"
            If Not Directory.Exists(d) Then
                Directory.CreateDirectory(d)
            End If
            Return d
        End Get
    End Property

    Public Shared ReadOnly Property SYSTEM_UTILITY_DIRECTORY As String
        Get
            Dim d = $"{RpaProject.SYSTEM_DIRECTORY}\util"
            If Not Directory.Exists(d) Then
                Directory.CreateDirectory(d)
            End If
            Return d
        End Get
    End Property

    Private _SystemUtilities As Dictionary(Of String, RpaUtility)
    Public Property SystemUtilities As Dictionary(Of String, RpaUtility)
        Get
            If Me._SystemUtilities Is Nothing Then
                Me._SystemUtilities = New Dictionary(Of String, RpaUtility)
            End If
            Return Me._SystemUtilities
        End Get
        Set(value As Dictionary(Of String, RpaUtility))
            Me._SystemUtilities = value
        End Set
    End Property

    Public Shared SYSTEM_JSON_FILENAME As String
    Public ReadOnly Property SystemJsonFileName As String
        Get
            If String.IsNullOrEmpty(RpaProject.SYSTEM_JSON_FILENAME) Then
                RpaProject.SYSTEM_JSON_FILENAME = RpaProject.SYSTEM_DIRECTORY & "\rpa_project.json"
                If Not File.Exists(RpaProject.SYSTEM_JSON_FILENAME) Then
                    Call Me.ModelSave(RpaProject.SYSTEM_JSON_FILENAME, Me)
                End If
            End If
            Return RpaProject.SYSTEM_JSON_FILENAME
        End Get
    End Property

    Public Shared ReadOnly Property SYSTEM_DLL_DIRECTORY As String
        Get
            Dim d = $"{RpaProject.SYSTEM_DIRECTORY}\dll"
            If Not Directory.Exists(d) Then
                Directory.CreateDirectory(d)
            End If
            Return d
        End Get
    End Property

    Private Shared _MyDirectory As String
    Public Property MyDirectory As String
        Get
            Return RpaProject._MyDirectory
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(RpaProject._MyDirectory) Then
                RpaProject._MyDirectory = value
                If Not Directory.Exists(RpaProject._MyDirectory) Then
                    Console.WriteLine("MyDirectory が存在しません")
                    Console.WriteLine("ファイル: " & Me.SystemJsonFileName & "の'MyDirectory' に任意のパスを書いてください")
                    Console.WriteLine("ファイルを保存した後、アプリケーションを再起動してください")
                End If
            End If
        End Set
    End Property

    Private _ProjectName As String
    Public Property ProjectName As String
        Get
            Return Me._ProjectName
        End Get
        Set(value As String)
            Me._ProjectName = value
            If Not String.IsNullOrEmpty(value) Then
                Call _CheckProject()
            End If
        End Set
    End Property

    Private _AliasDictionary As Dictionary(Of String, String)
    Public Property AliasDictionary As Dictionary(Of String, String)
        Get
            If Me._AliasDictionary Is Nothing Then
                Me._AliasDictionary = New Dictionary(Of String, String)
            End If
            Return Me._AliasDictionary
        End Get
        Set(value As Dictionary(Of String, String))
            Me._AliasDictionary = value
        End Set
    End Property

    Private _ProjectAlias As String
    Public ReadOnly Property ProjectAlias As String
        Get
            If String.IsNullOrEmpty(Me._ProjectAlias) Then
                If String.IsNullOrEmpty(Me.ProjectName) Then
                    Me._ProjectAlias = vbNullString
                Else
                    If Me.AliasDictionary.ContainsKey(Me.ProjectName) Then
                        Me._ProjectAlias = Me.AliasDictionary(Me.ProjectName)
                    Else
                        Me._ProjectAlias = Me.ProjectName
                    End If
                End If
            End If
            Return Me._ProjectAlias
        End Get
    End Property

    Private Shared _ServerDirectory As String
    Public Shared Property ServerDirectory As String
        Get
            Return RpaProject._ServerDirectory
        End Get
        Set(value As String)
            If Not String.IsNullOrEmpty(value) Then
                RpaProject._ServerDirectory = value
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
                        obj2 = obj.ModelLoad(Me.RootProjectJsonFileName)
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
            Return Me.RootProjectDirectory & "\ignore"
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
                    Me._RootProjectIgnoreList = _GetFileLines(Me.RootProjectIgnoreFileName)
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
    '                        pack = jh.ModelLoad(Of RpaPackage)(d & "\RpaPackage.json")
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
    Public ReadOnly Property IsMyProjectExists As Boolean
        Get
            If Directory.Exists(Me.MyProjectDirectory) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property MyProjectDirectory As String
        Get
            Return $"{Me.MyDirectory}\{Me.ProjectAlias}"
        End Get
    End Property

    'Public ReadOnly Property MyProjectSysDirectory As String
    '    Get
    '        Return Me.MyProjectDirectory & "\sys"
    '    End Get
    'End Property

    'Private _MyProjectWorkDirectory As String
    'Public ReadOnly Property MyProjectWorkDirectory As String
    '    Get
    '        If String.IsNullOrEmpty(Me._MyProjectWorkDirectory) Then
    '            Me._MyProjectWorkDirectory = Me.MyProjectDirectory & "\work"
    '            If Directory.Exists(Me._MyProjectDirectory) Then
    '                If Not Directory.Exists(Me._MyProjectWorkDirectory) Then
    '                    Directory.CreateDirectory(Me._MyProjectWorkDirectory)
    '                End If
    '            End If
    '        End If
    '        Return Me._MyProjectWorkDirectory
    '    End Get
    'End Property

    ' 現状不要 2020-12-11
    'Private _MyProjectScriptDirectory As String
    'Public ReadOnly Property MyProjectScriptDirectory As String
    '    Get
    '        If String.IsNullOrEmpty(Me._MyProjectScriptDirectory) Then
    '            Me._MyProjectScriptDirectory = Me.MyProjectDirectory & "\script"
    '            If Directory.Exists(Me._MyProjectScriptDirectory) Then
    '                Directory.CreateDirectory(Me._MyProjectScriptDirectory)
    '            End If
    '        End If
    '        Return Me._MyProjectScriptDirectory
    '    End Get
    'End Property

    ' 成果物保管用
    'Private _MyProjectBackupDirectory As String
    'Public ReadOnly Property MyProjectBackupDirectory As String
    '    Get
    '        If String.IsNullOrEmpty(Me._MyProjectBackupDirectory) Then
    '            Me._MyProjectBackupDirectory = Me.MyProjectDirectory & "\backup"
    '            If Directory.Exists(Me._MyProjectDirectory) Then
    '                If Not Directory.Exists(Me._MyProjectBackupDirectory) Then
    '                    Directory.CreateDirectory(Me._MyProjectBackupDirectory)
    '                End If
    '            End If
    '        End If
    '        Return Me._MyProjectBackupDirectory
    '    End Get
    'End Property

    Private _MyProjectJsonFileName As String
    Public ReadOnly Property MyProjectJsonFileName As String
        Get
            If String.IsNullOrEmpty(Me._MyProjectJsonFileName) Then
                Me._MyProjectJsonFileName = $"{Me.MyProjectDirectory}\rpa_project.json"
            End If
            Return Me._MyProjectJsonFileName
        End Get
    End Property

    Private Property _MyProjectObject As Object
    <JsonIgnore>
    Public Property MyProjectObject As Object
        Get
            If Me._MyProjectObject Is Nothing Then
                Dim obj = RpaCodes.RpaObject(Me)
                Dim obj2 = Nothing
                If obj Is Nothing Then
                    Me._MyProjectObject = Nothing
                Else
                    If File.Exists(Me.MyProjectJsonFileName) Then
                        obj2 = obj.ModelLoad(Me.MyProjectJsonFileName)
                    End If
                    If obj2 Is Nothing Then
                        Me._MyProjectObject = obj
                    Else
                        Me._MyProjectObject = obj2
                    End If
                End If
            End If
            Return Me._MyProjectObject
        End Get
        Set(value As Object)
            Me._MyProjectObject = value
        End Set
    End Property

    Public ReadOnly Property MyProjectIgnoreFileName As String
        Get
            Return Me.MyProjectDirectory & "\ignore"
        End Get
    End Property

    Private _MyProjectIgnoreList As List(Of String)
    <JsonIgnore>
    Public Property MyProjectIgnoreList As List(Of String)
        Get
            If Me._MyProjectIgnoreList Is Nothing Then
                Me._MyProjectIgnoreList = New List(Of String)
                If File.Exists(Me.MyProjectIgnoreFileName) Then
                    Me._MyProjectIgnoreList = _GetFileLines(Me.MyProjectIgnoreFileName)
                End If
            End If
            Return Me._MyProjectIgnoreList
        End Get
        Set(value As List(Of String))
            Me._MyProjectIgnoreList = value
        End Set
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
    '                        pack = jh.ModelLoad(Of RpaPackage)(f)
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

    Private _PrinterName As String
    Public Property PrinterName As String
        Get
            Return Me._PrinterName
        End Get
        Set(value As String)
            Me._PrinterName = value
        End Set
    End Property

    Private _UsePrinterName As String
    <JsonIgnore>
    Public Property UsePrinterName As String
        Get
            Return Me._UsePrinterName
        End Get
        Set(value As String)
            Me._UsePrinterName = value
        End Set
    End Property


    Private Sub _CheckProject()
        Dim rpack As RpaPackage = Nothing
        Dim mpack As RpaPackage = Nothing
        Console.WriteLine("プロジェクトの検査中...")
        Console.WriteLine("プロジェクト名                 : " & Me.ProjectName)
        Console.WriteLine("プロジェクト別名               : " & Me.ProjectAlias)
        Console.WriteLine("ルートプロジェクトディレクトリ : " & Me.RootProjectDirectory)
        If Not Me.IsRootProjectExists Then
            Console.WriteLine("(警告) ルートプロジェクトディレクトリは存在しません")
        End If
        Console.WriteLine("ユーザプロジェクトディレクトリ : " & Me.MyProjectDirectory)
        If Not Me.IsMyProjectExists Then
            Console.WriteLine("       ユーザプロジェクトディレクトリは存在しません")
        End If

        'Dim installed = False
        'Console.WriteLine("アップデートパッケージを検索しています...")
        'For Each rp In Me.RootProjectUpdatePackages
        '    installed = False
        '    For Each mp In Me.MyProjectUpdatedPackages
        '        If rp.Name = mp.Name Then
        '            installed = True
        '            Exit For
        '        End If
        '    Next
        '    If Not installed Then
        '        Console.WriteLine("パッケージ : " & rp.Name & " をインストールしていません")
        '    End If
        'Next
        Console.WriteLine("プロジェクトの検査完了")
    End Sub

    'Public Sub InvokeOutlookMacro(ByVal method As String, ParamArray args As Object())
    '    Dim olapp, exbooks, exbook
    '    Dim macrofile As String
    '    Try
    '        olapp = CreateObject("Outlook.Application")
    '        Try
    '            exbooks = olapp.Workbooks
    '            Try
    '                exbook = exbooks.Open(Me.SystemMacroFileName, 0)
    '                macrofile = Path.GetFileName(Me.SystemMacroFileName)
    '                Select Case args.Length
    '                    Case 1
    '                        olapp.Run($"{macrofile}!{method}", args)
    '                    Case 2
    '                        olapp.Run($"{macrofile}!{method}", args(0), args(1))
    '                End Select

    '            Finally
    '                If exbook IsNot Nothing Then
    '                    exbook.Close()
    '                End If
    '                Marshal.ReleaseComObject(exbook)
    '            End Try
    '        Finally
    '            Marshal.ReleaseComObject(exbooks)
    '        End Try
    '    Finally
    '        If olapp IsNot Nothing Then
    '            olapp.Quit()
    '        End If
    '        Marshal.ReleaseComObject(olapp)
    '    End Try
    'End Sub

    'Public Sub InvokeMacro(ByVal method As String, args() As Object)
    '    Dim exapp, exbooks, exbook
    '    Dim macrofile As String
    '    Try
    '        exapp = CreateObject("Excel.Application")
    '        Try
    '            exbooks = exapp.Workbooks
    '            Try
    '                exbook = exbooks.Open(Me.SystemMacroFileName, 0)
    '                macrofile = Path.GetFileName(Me.SystemMacroFileName)
    '                exapp.Run($"{macrofile}!{method}", args)
    '            Finally
    '                If exbook IsNot Nothing Then
    '                    exbook.Close()
    '                End If
    '                Marshal.ReleaseComObject(exbook)
    '            End Try
    '        Finally
    '            Marshal.ReleaseComObject(exbooks)
    '        End Try
    '    Finally
    '        If exapp IsNot Nothing Then
    '            exapp.Quit()
    '        End If
    '        Marshal.ReleaseComObject(exapp)
    '    End Try
    'End Sub

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

    Private Function _GetFileLines(ByVal f As String) As List(Of String)
        Dim txt As String
        Dim olines As String()
        Dim nlines = New List(Of String)

        Dim sr = New StreamReader(f, Encoding.GetEncoding(MYENCODING))

        txt = sr.ReadToEnd()
        olines = txt.Split(vbCrLf)

        For Each line In olines
            line = IIf(line.Contains(vbCr), line.Trim(vbCr), line)
            line = IIf(line.Contains(vbLf), line.Trim(vbLf), line)
            nlines.Add(line)
        Next

        sr.Close()
        sr.Dispose()

        Return nlines
    End Function

    Public Sub New()
    End Sub
End Class
