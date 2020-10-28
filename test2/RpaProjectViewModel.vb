Imports test2
Imports System.Collections.ObjectModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Forms
Imports System.IO

Public Class RpaProjectViewModel : Inherits BaseViewModel2
    Private Const SHIFT_JIS As String = "Shift-JIS"

    Private _RootDirectoryName As String
    Public Property RootDirectoryName As String
        Get
            If Me._RootDirectoryName Is Nothing Then
                Me._RootDirectoryName = "ルートディレクトリを入力してください"
            End If
            Return Me._RootDirectoryName
        End Get
        Set(value As String)
            Me._RootDirectoryName = value
            RaisePropertyChanged("RootDirectoryName")
            AppInfo.ProjectInfo.Model.Data.RootDirectoryName = value
            Call _CheckLaunchCommandEnabled()
        End Set
    End Property

    Private _UserDirectoryName As String
    Public Property UserDirectoryName As String
        Get
            If Me._UserDirectoryName Is Nothing Then
                Me._UserDirectoryName = "ユーザーディレクトリを入力してください"
            End If
            Return Me._UserDirectoryName
        End Get
        Set(value As String)
            Me._UserDirectoryName = value
            RaisePropertyChanged("UserDirectoryName")
            AppInfo.ProjectInfo.Model.Data.UserDirectoryName = value
            Call _CheckLaunchCommandEnabled()
        End Set
    End Property

    Private _SystemDirectoryName As String
    Public Property SystemDirectoryName As String
        Get
            If Me._SystemDirectoryName Is Nothing Then
                Me._SystemDirectoryName = "システムディレクトリを入力してください"
            End If
            Return Me._SystemDirectoryName
        End Get
        Set(value As String)
            Me._SystemDirectoryName = value
            RaisePropertyChanged("SystemDirectoryName")
            AppInfo.ProjectInfo.Model.Data.SystemDirectoryName = value
            Call _CheckLaunchCommandEnabled()
        End Set
    End Property

    Private _PythonPathName As String
    Public Property PythonPathName As String
        Get
            If Me._PythonPathName Is Nothing Then
                Me._PythonPathName = "Input Python Path"
            End If
            Return Me._PythonPathName
        End Get
        Set(value As String)
            Me._PythonPathName = value
            RaisePropertyChanged("PythonPathName")
            AppInfo.ProjectInfo.Model.Data.PythonPathName = value
            Call _CheckLaunchCommandEnabled()
        End Set
    End Property

    Private _ErrorMessage As String
    Public Property ErrorMessage As String
        Get
            Return Me._ErrorMessage
        End Get
        Set(value As String)
            Me._ErrorMessage = value
            RaisePropertyChanged("ErrorMessage")
        End Set
    End Property

    Private _Rpas As ObservableCollection(Of RpaModel)
    <JsonIgnore>
    Public Property Rpas As ObservableCollection(Of RpaModel)
        Get
            If Me._Rpas Is Nothing Then
                Me._Rpas = New ObservableCollection(Of RpaModel)
            End If
            Return Me._Rpas
        End Get
        Set(value As ObservableCollection(Of RpaModel))
            Me._Rpas = value
        End Set
    End Property

    ' コマンドプロパティ（フォルダ選択）
    '---------------------------------------------------------------------------------------------'
    Private _FolderBrowserDialogCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property FolderBrowserDialogCommand As ICommand
        Get
            If Me._FolderBrowserDialogCommand Is Nothing Then
                Me._FolderBrowserDialogCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _FolderBrowserDialogCommandExecute,
                    .CanExecuteHandler = AddressOf _FolderBrowserDialogCommandCanExecute
                }
                Return Me._FolderBrowserDialogCommand
            Else
                Return Me._FolderBrowserDialogCommand
            End If
        End Get
    End Property

    Private Sub _CheckFolderBrowserDialogCommandEnabled()
        Dim b As Boolean : b = True
        Me._FolderBrowserDialogCommandEnableFlag = b
    End Sub

    Private __FolderBrowserDialogCommandEnableFlag As Boolean
    Private Property _FolderBrowserDialogCommandEnableFlag As Boolean
        Get
            Return Me.__FolderBrowserDialogCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__FolderBrowserDialogCommandEnableFlag = value
            RaisePropertyChanged("_FolderBrowserDialogCommandEnableFlag")
            CType(FolderBrowserDialogCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _FolderBrowserDialogCommandExecute(ByVal parameter As Object)
        Dim dsc As String
        Dim fbd As FolderBrowserDialog
        Dim f = CType(parameter, String)
        Dim b = True

        Select Case f
            Case Me.RootDirectoryName
                dsc = "ルートディレクトリの指定"
            Case Me.SystemDirectoryName
                dsc = "システムディレクトリの指定"
                b = False
            Case Me.UserDirectoryName
                dsc = "ユーザーディレクトリの指定"
            Case Me.PythonPathName
                dsc = "Python Path Setting"
        End Select

        fbd = New FolderBrowserDialog With {
            .Description = dsc,
            .SelectedPath = $"{f}",
            .ShowNewFolderButton = False
        }
        If fbd.ShowDialog() = DialogResult.OK Then
            Select Case f
                Case Me.RootDirectoryName
                    Me.RootDirectoryName = fbd.SelectedPath
                Case Me.SystemDirectoryName
                    Me.SystemDirectoryName = fbd.SelectedPath
                Case Me.UserDirectoryName
                    Me.UserDirectoryName = fbd.SelectedPath
                Case Me.PythonPathName
                    Me.PythonPathName = fbd.SelectedPath
            End Select
        End If
    End Sub

    Private Function _FolderBrowserDialogCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._FolderBrowserDialogCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    ' コマンドプロパティ（プロジェクト作成）
    '---------------------------------------------------------------------------------------------'
    Private _LaunchCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property LaunchCommand As ICommand
        Get
            If Me._LaunchCommand Is Nothing Then
                Me._LaunchCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _LaunchCommandExecute,
                    .CanExecuteHandler = AddressOf _LaunchCommandCanExecute
                }
                Return Me._LaunchCommand
            Else
                Return Me._LaunchCommand
            End If
        End Get
    End Property

    Private Sub _CheckLaunchCommandEnabled()
        Dim b As Boolean : b = True
        Dim dq = """"
        Dim msg = vbNullString
        If Not Directory.Exists(Me.PythonPathName) Then
            b = False
            msg = $"Error: Python Path{dq}{Me.PythonPathName}{dq} Is Not Exist"
        End If
        If Not Directory.Exists(Me.RootDirectoryName) Then
            b = False
            msg = $"Error: ユーザーディレクトリ{dq}{Me.UserDirectoryName}{dq} は存在しません"
        End If
        If Not Directory.Exists(Me.RootDirectoryName) Then
            b = False
            msg = $"Error: ルートディレクトリ{dq}{Me.RootDirectoryName}{dq} は存在しません"
        End If
        If b Then
            If Directory.Exists(Me.SystemDirectoryName) Then
                b = False
                msg = $"Error: システムディレクトリの同名ディレクトリが既に存在します"
            End If
        End If
        Me._LaunchCommandEnableFlag = b
        Me.ErrorMessage = msg
    End Sub

    Private __LaunchCommandEnableFlag As Boolean
    Private Property _LaunchCommandEnableFlag As Boolean
        Get
            Return Me.__LaunchCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__LaunchCommandEnableFlag = value
            RaisePropertyChanged("_LaunchCommandEnableFlag")
            CType(LaunchCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _LaunchCommandExecute(ByVal parameter As Object)
        Dim b As Boolean
        Call _CheckLaunchCommandEnabled()
        If _LaunchCommandEnableFlag Then
            Directory.CreateDirectory(Me.SystemDirectoryName)
            With AppInfo.ProjectInfo.Model.Data
                Directory.CreateDirectory(.SysDirectoryName)
                b = _MakeFile(.MyDirFileName, Me.UserDirectoryName)
                b = _MakeFile(.MyPythonFileName, Me.PythonPathName)
                b = _MakeFile(.RpaProjectFileName, vbNullString)
                b = _FilesImportRecursive(.RootSysDirectoryName, .SysDirectoryName)
            End With
        End If
        Call _CheckLaunchCommandEnabled()
    End Sub

    Private Function _LaunchCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._LaunchCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------    

    Private Function _FilesImportRecursive(ByVal src As String, ByVal dst As String) As Boolean
        Dim b = False
        Dim sdi = New DirectoryInfo(src)
        Try
            For Each f In sdi.GetFiles
                f.CopyTo($"{dst}\{f.Name}")
                b = True
            Next
            For Each d In sdi.GetDirectories
                b = _FilesImportRecursive(d.FullName, dst)
            Next
        Catch ex As Exception
            b = False
        Finally
            _FilesImportRecursive = False
        End Try
    End Function


    Private Function _MakeFile(ByVal f As String, ByVal txt As String) As Boolean
        Dim b = False
        Dim sw As System.IO.StreamWriter
        Try
            sw = New System.IO.StreamWriter(
                f, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
            )
            sw.Write(txt)
            b = True
        Catch ex As Exception
            b = False
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
            _MakeFile = b
        End Try
    End Function

    Public Sub _ViewInitializing()
        With AppInfo.ProjectInfo.Model.Data
            Me.RootDirectoryName = .RootDirectoryName
            Me.UserDirectoryName = .UserDirectoryName
            Me.SystemDirectoryName = .SystemDirectoryName
            Me.PythonPathName = .PythonPathName
            Me.Rpas = .Rpas
        End With
    End Sub


    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        InitializeHandler = AddressOf _ViewInitializing
        CheckCommandEnabledHandler _
            = [Delegate].Combine(
                New Action(AddressOf _CheckFolderBrowserDialogCommandEnabled),
                New Action(AddressOf _CheckLaunchCommandEnabled)
            )
        Call BaseInitialize(app, vm)
    End Sub
End Class
