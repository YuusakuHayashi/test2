Imports System.IO
Imports System.Text

Public Class RpaProject : Inherits JsonHandler(Of RpaProject)
    Private Const SHIFT_JIS As String = "Shift-JIS"

    Public Shared ROOT_DIRECTORY As String = "\\Coral\受託_計算Ｃ本部\連絡・管理・トラ・教育他\教育\RPAプロジェクト"
    Public Shared SYSTEM_DIRECTORY As String = System.Environment.GetEnvironmentVariable("USERPROFILE") & "\rpa_project"
    Public Shared SYSTEM_WORK_DIRECTORY As String = SYSTEM_DIRECTORY & "\work"
    Public Shared SYSTEM_SYSTEM_DIRECTORY As String = SYSTEM_DIRECTORY & "\sys"
    Public Shared SYSTEM_JSON_FILENAME As String = SYSTEM_DIRECTORY & "\rpa_project.json"
    Public Shared SYSTEM_WORK_JSON_FILENAME As String = SYSTEM_WORK_DIRECTORY & "\rpa_project.json"
    Public Shared SYSTEM_SCRIPT_DIRECTORY As String = SYSTEM_DIRECTORY & "\script"
    Public Shared MYDIRECTORY_FILE As String = SYSTEM_DIRECTORY & "\mydir"

    Private _MyDirectory As String
    Public Property MyDirectory As String
        Get
            Return Me._MyDirectory
        End Get
        Set(value As String)
            Me._MyDirectory = value
        End Set
    End Property

    Private _ProjectName As String
    Public Property ProjectName As String
        Get
            Return Me._ProjectName
        End Get
        Set(value As String)
            Me._ProjectName = value
            Call _GetAlias(value)
            Console.WriteLine("プロジェクト選択 => " & Me.ProjectName & "(" & Me.ProjectAlias & ")")
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
    Public Property ProjectAlias As String
        Get
            Return Me._ProjectAlias
        End Get
        Set(value As String)
            Me._ProjectAlias = value
        End Set
    End Property

    Private _ServerDirectory As String
    Public Property ServerDirectory As String
        Get
            Return Me._ServerDirectory
        End Get
        Set(value As String)
            Me._ServerDirectory = value
        End Set
    End Property

    ' Root Directory 関係
    '-----------------------------------------------------------------------------------'
    Public ReadOnly Property IsRootProjectExists As Boolean
        Get
            If Directory.Exists(Me.RootProjectDirectory) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property RootProjectDirectory As String
        Get
            Return RpaProject.ROOT_DIRECTORY & "\" & Me.ProjectName
        End Get
    End Property

    Public ReadOnly Property RootProjectJsonFileName As String
        Get
            Return Me.RootProjectDirectory & "\rpa_project.json"
        End Get
    End Property

    Private Property _RootProjectObject As Object
    Public Property RootProjectObject As Object
        Get
            If Me._RootProjectObject Is Nothing Then
                Dim obj = RpaCodes.RpaObject(Me.ProjectName)
                Dim obj2 = Nothing
                If Directory.Exists(Me.RootProjectJsonFileName) Then
                    obj2 = obj.ModelLoad(Me.RootProjectJsonFileName)
                End If
                If obj2 IsNot Nothing Then
                    Me._RootProjectObject = obj2
                Else
                    Me._RootProjectObject = obj
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

    Private _RootProjectIgnoreList As List(Of String)
    Public ReadOnly Property RootProjectIgnoreList As List(Of String)
        Get
            Dim lines As String()
            If Me._RootProjectIgnoreList Is Nothing Then
                Me._RootProjectIgnoreList = New List(Of String)
                If File.Exists(Me.RootProjectIgnoreFileName) Then
                    lines = _GetFileLines(Me.RootProjectIgnoreFileName)
                    For Each line In lines
                        Me._RootProjectIgnoreList.Add(line)
                    Next
                End If
            End If
            Return Me._RootProjectIgnoreList
        End Get
    End Property

    Public ReadOnly Property RootProjectUpdateDirectory As String
        Get
            Return Me.RootProjectDirectory & "\updates"
        End Get
    End Property

    Public ReadOnly Property RootProjectUpdatePackages As String()
        Get
            If Directory.Exists(Me.RootProjectUpdateDirectory) Then
                Return Directory.GetDirectories(Me.RootProjectUpdateDirectory)
            Else
                Return Nothing
            End If
        End Get
    End Property
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
            Return Me.MyDirectory & "\" & Me.ProjectAlias
        End Get
    End Property

    Public ReadOnly Property MyProjectSysDirectory As String
        Get
            Return Me.MyProjectDirectory & "\sys"
        End Get
    End Property

    Public ReadOnly Property MyProjectWorkDirectory As String
        Get
            Return Me.MyProjectDirectory & "\work"
        End Get
    End Property

    Public ReadOnly Property MyProjectScriptDirectory As String
        Get
            Return Me.MyProjectDirectory & "\script"
        End Get
    End Property

    Public ReadOnly Property MyProjectJsonFileName As String
        Get
            Return Me.MyProjectDirectory & "\rpa_project.json"
        End Get
    End Property

    Private Property _MyProjectObject As Object
    Public Property MyProjectObject As Object
        Get
            If Me._MyProjectObject Is Nothing Then
                Dim obj = RpaCodes.RpaObject(Me.ProjectName)
                Dim obj2 = Nothing
                If Directory.Exists(Me.MyProjectJsonFileName) Then
                    obj2 = obj.ModelLoad(Me.MyProjectJsonFileName)
                End If
                If obj2 IsNot Nothing Then
                    Me._MyProjectObject = obj2
                Else
                    Me._MyProjectObject = obj
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
    Public ReadOnly Property MyProjectIgnoreList As List(Of String)
        Get
            Dim lines As String()
            If Me._MyProjectIgnoreList Is Nothing Then
                Me._MyProjectIgnoreList = New List(Of String)
                If File.Exists(Me.MyProjectIgnoreFileName) Then
                    lines = _GetFileLines(Me.MyProjectIgnoreFileName)
                    For Each line In lines
                        Me._MyProjectIgnoreList.Add(line)
                    Next
                End If
            End If
            Return Me._MyProjectIgnoreList
        End Get
    End Property

    Public ReadOnly Property MyProjectUpdatedFileName As String
        Get
            Return Me.MyProjectDirectory & "\updated"
        End Get
    End Property

    Private _MyProjectUpdatedPackages As List(Of String)
    Public ReadOnly Property MyProjectUpdatedPackages As List(Of String)
        Get
            Dim lines As String()
            If Me._MyProjectUpdatedPackages Is Nothing Then
                Me._MyProjectUpdatedPackages = New List(Of String)
                If File.Exists(Me.MyProjectUpdatedFileName) Then
                    lines = _GetFileLines(Me.MyProjectUpdatedFileName)
                    For Each line In lines
                        Me._MyProjectUpdatedPackages.Add(line)
                    Next
                End If
            End If
            Return Me._MyProjectUpdatedPackages
        End Get
    End Property
    '-----------------------------------------------------------------------------------'

    Private Sub _GetAlias(ByVal project As String)
        If Not String.IsNullOrEmpty(project) Then
            If Me.AliasDictionary.ContainsKey(project) Then
                Me.ProjectAlias = Me.AliasDictionary(project)
            Else
                Me.ProjectAlias = project
            End If
        End If
    End Sub

    Public Sub CheckSystemConstitution()
        Call _CheckMyDirectory()
        Call _CheckServerDirectory()
    End Sub

    ' ServerDirectory指定がある場合、ROOT_DIRECTORYを切り替える
    Private Sub _CheckServerDirectory()
        If Directory.Exists(Me.ServerDirectory) Then
            Console.WriteLine("ServerDirectory: " & Me.ServerDirectory)
            Console.WriteLine("Switch ROOT_DIRECTORY: " & RpaProject.ROOT_DIRECTORY)
            RpaProject.ROOT_DIRECTORY = Me.ServerDirectory
            Console.WriteLine(" => To ROOT_DIRECTORY: " & RpaProject.ROOT_DIRECTORY)
        End If
    End Sub

    Private Sub _CheckMyDirectory()
        Dim rpa As RpaProject
        Dim answer As String

        Console.WriteLine("MyDirectory Checking...")

        Do Until Directory.Exists(Me.MyDirectory)
            File.Copy(RpaProject.SYSTEM_JSON_FILENAME, RpaProject.SYSTEM_WORK_JSON_FILENAME, True)

            Console.WriteLine("MyDirectory が存在しません")
            Console.WriteLine("ファイル: " & RpaProject.SYSTEM_WORK_JSON_FILENAME & "の'MyDirectory' に任意のパスを書いてください")
            Console.WriteLine("書いた後ファイルを保存し、[Enter]を入力してください")
            Console.ReadLine()

            rpa = Me.ModelLoad(RpaProject.SYSTEM_WORK_JSON_FILENAME)
            Me.MyDirectory = rpa.MyDirectory

            Console.WriteLine("よろしいですか？(y/n) MyDirectory => " & Me.MyDirectory)
            answer = Console.ReadLine()
            If Not answer = "y" Then
                Me.MyDirectory = vbNullString
            End If

            File.Copy(RpaProject.SYSTEM_WORK_JSON_FILENAME, RpaProject.SYSTEM_JSON_FILENAME, True)
        Loop

        Console.WriteLine("MyDirectory Check OK!")
    End Sub

    Private Function _GetFileLines(ByVal f As String) As String()
        Dim txt As String
        Dim lines As String()

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        Dim sr = New StreamReader(f, Encoding.GetEncoding(SHIFT_JIS))

        txt = sr.ReadToEnd()
        lines = txt.Split(vbCrLf)

        sr.Close()
        sr.Dispose()

        Return lines
    End Function

    Public Sub MakeStaff()

        Dim mkfil As Action(Of String)
        mkfil = Sub(ByVal f As String)
                    If Not File.Exists(f) Then
                        File.Create(f)
                    End If
                End Sub
        Dim mkdir As Action(Of String)
        mkdir = Sub(ByVal d As String)
                    If Not Directory.Exists(d) Then
                        Directory.CreateDirectory(d)
                    End If
                End Sub

        Call mkdir(RpaProject.SYSTEM_DIRECTORY)
        Call mkdir(RpaProject.SYSTEM_SYSTEM_DIRECTORY)
        Call mkdir(RpaProject.SYSTEM_WORK_DIRECTORY)
        Call mkfil(RpaProject.SYSTEM_JSON_FILENAME)
    End Sub

    Public Sub New()
    End Sub
End Class
