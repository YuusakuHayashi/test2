Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Forms
Imports System.ComponentModel

'Public Class RpaProject : Inherits RpaCui.JsonHandler(Of RpaProject)
Public Class IntranetClientServerProject
    Inherits RpaProjectBase(Of IntranetClientServerProject)

    Private Const SHIFT_JIS As String = "Shift-JIS"
    Private Const UTF8 As String = "utf-8"

    Private Const ARCHITECTURE_NAME As String = "IntranetClientServer"

    Public Overrides ReadOnly Property GuideString As String
        Get
            Dim upck As String = vbNullString
            If Me.UpdateAvaileble Then
                upck = "(!)"
            End If
            Return $"{upck}{Me.ProjectName}\{Me.RobotAlias}>"
        End Get
    End Property

    ' Update 関係
    '-----------------------------------------------------------------------------------'
    ' Updateフラグ
    Private _UpdateAvaileble As Boolean
    <JsonIgnore>
    Public Property UpdateAvaileble As Boolean
        Get
            Return Me._UpdateAvaileble
        End Get
        Set(value As Boolean)
            Me._UpdateAvaileble = value
        End Set
    End Property

    ' 自動更新タスク(CheckUpdateAvailable()) を起動するためのチェック
    ' 条件を満たした初回時のみTrueにする
    Private FirstUpdateAvailable As Boolean
    Private ReadOnly Property IsUpdateAvailable() As Boolean
        Get
            Dim b As Boolean = False
            If File.Exists(Me.RootRobotsUpdateFile) And File.Exists(Me.SystemRobotsUpdateFile) Then
                If Not FirstUpdateAvailable Then
                    FirstUpdateAvailable = True
                    b = True
                End If
            End If
            Return b
        End Get
    End Property
    '-----------------------------------------------------------------------------------'

    Private _ReleaseRobotDirectory As String
    Public Property ReleaseRobotDirectory As String
        Get
            Return Me._ReleaseRobotDirectory
        End Get
        Set(value As String)
            Me._ReleaseRobotDirectory = value
        End Set
    End Property

    ' Root Directory 関係
    '-----------------------------------------------------------------------------------'
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

    <JsonIgnore>
    Public ReadOnly Property RootDllDirectory As String
        Get
            Dim [dir] As String = vbNullString
            If Directory.Exists(Me.RootDirectory) Then
                [dir] = $"{Me.RootDirectory}\dll"
                If Not Directory.Exists([dir]) Then
                    Directory.CreateDirectory([dir])
                End If
            End If
            Return [dir]
        End Get
    End Property

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

    ' このファイルがRootRobotになければ、AttachRobotすることができない
    <JsonIgnore>
    Public ReadOnly Property RootRobotReadMeFileName As String
        Get
            Dim [fil] As String = vbNullString
            If Directory.Exists(Me.RootRobotDirectory) Then
                [fil] = $"{Me.RootRobotDirectory}\readme"
            End If
            Return [fil]
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

    <JsonIgnore>
    Public ReadOnly Property RootRobotsUpdateFile As String
        Get
            Dim fil As String = vbNullString
            Dim jh As New RpaCui.JsonHandler(Of List(Of RpaUpdater))
            If Not String.IsNullOrEmpty(Me.RootDirectory) Then
                fil = $"{Me.RootDirectory}\robotsupdate"
                If Not File.Exists(fil) Then
                    Call jh.Save(fil, (New List(Of RpaUpdater)))
                End If
            End If
            Return fil
        End Get
    End Property

    Public ReadOnly Property MyRobotIgnoreFileName As String
        Get
            Dim [fil] As String = vbNullString
            If Directory.Exists(Me.MyRobotDirectory) Then
                [fil] = $"{Me.MyRobotDirectory}\ignore"
            End If
            Return [fil]
        End Get
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

    <JsonIgnore>
    Public ReadOnly Property SystemRobotsUpdateFile As String
        Get
            Dim jh As New RpaCui.JsonHandler(Of List(Of RpaUpdater))
            Dim fil As String = $"{RpaCui.SystemDirectory}\robotsupdate"
            If Not File.Exists(fil) Then
                Call jh.Save(fil, (New List(Of RpaUpdater)))
            End If
            Return fil
        End Get
    End Property

    Public Overrides Function SwitchRobot(name As String) As Integer
        Dim i As Integer = MyBase.SwitchRobot(name)
        Me.RootRobotObject = Nothing

        Return 0
    End Function

    Private Sub _CheckProject()
    End Sub

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

    Private Sub CheckUpdateAvailable(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If Me.IsUpdateAvailable Then
            Call _CheckUpdateAvailable()
        End If
    End Sub

    Private Async Sub _CheckUpdateAvailable()
        Dim jh As New RpaCui.JsonHandler(Of RpaUpdater)
        Dim rrus As List(Of RpaUpdater)
        Dim srus As List(Of RpaUpdater)
        Dim t As Task = Task.Run(
            Sub()
                Do
                    rrus = jh.Load(Of List(Of RpaUpdater))(Me.RootRobotsUpdateFile)
                    srus = jh.Load(Of List(Of RpaUpdater))(Me.SystemRobotsUpdateFile)

                    If rrus.Count > 0 Then
                        rrus.Sort(
                            Function(before, after)
                                Return (before.ReleaseDate < after.ReleaseDate)
                            End Function
                        )
                    End If

                    If srus.Count > 0 Then
                        srus.Sort(
                            Function(before, after)
                                Return (before.ReleaseDate < after.ReleaseDate)
                            End Function
                        )
                    End If

                    If rrus.Count > 0 Then
                        If srus.Count = 0 Then
                            Me.UpdateAvaileble = True
                        Else
                            If rrus.Last.ReleaseDate > srus.Last.ReleaseDate Then
                                Me.UpdateAvaileble = True
                            End If
                        End If
                    End If

                    Threading.Thread.Sleep(60000)
                Loop Until Me.UpdateAvaileble
            End Sub
        )
        Await t
    End Sub

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
        AddHandler Me.PropertyChanged, AddressOf CheckUpdateAvailable
    End Sub


    Protected Overrides Sub Finalize()
        RemoveHandler Me.PropertyChanged, AddressOf CreateChangedFile
        MyBase.Finalize()
    End Sub
End Class
