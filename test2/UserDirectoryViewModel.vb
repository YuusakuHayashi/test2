Imports System.Windows.Forms
Imports System.Collections.ObjectModel
Imports System.IO

Public Class UserDirectoryViewModel
    Inherits BaseViewModel2
    'Inherits BaseViewModel2(Of Object)

    Private Const _PROJECT_LOAD_FAILED As String = "プロジェクトのロードに失敗しました"
    Private Const _ISNOT_PROJECT_DIRECTORY As String = "このフォルダはプロジェクトディレクトリではありません"
    Private Const _DIRECTORY_ALREADY_EXIST As String = "このフォルダは既に存在しています"

    Private _UserDirectoryName As String
    Public Property UserDirectoryName As String
        Get
            Return _UserDirectoryName
        End Get
        Set(value As String)
            _UserDirectoryName = value
            RaisePropertyChanged("UserDirectoryName")
            AppInfo.ProjectInfo.RootDirectoryName = value
            Call _UpdateProjectDirectoryName()
            Call _CheckAddProjectCommandEnabled()
        End Set
    End Property

    Private _ProjectName As String
    Public Property ProjectName As String
        Get
            Return _ProjectName
        End Get
        Set(value As String)
            _ProjectName = value
            AppInfo.ProjectInfo.Name = value
            RaisePropertyChanged("ProjectName")
            Call _UpdateProjectDirectoryName()
            Call _CheckAddProjectCommandEnabled()
        End Set
    End Property

    Private _ProjectDirectoryName As String
    Public Property ProjectDirectoryName As String
        Get
            Return _ProjectDirectoryName
        End Get
        Set(value As String)
            _ProjectDirectoryName = value
            AppInfo.ProjectInfo.DirectoryName = value
            RaisePropertyChanged("ProjectDirectoryName")
        End Set
    End Property

    ' 
    Private _CurrentProjects As ObservableCollection(Of ProjectInfoModel)
    Public Property CurrentProjects As ObservableCollection(Of ProjectInfoModel)
        Get
            If _CurrentProjects Is Nothing Then
                _CurrentProjects = New ObservableCollection(Of ProjectInfoModel)
            End If
            Return _CurrentProjects
        End Get
        Set(value As ObservableCollection(Of ProjectInfoModel))
            _CurrentProjects = value
            AppInfo.CurrentProjects = value
            RaisePropertyChanged("CurrentProjects")
        End Set
    End Property

    Private _FixedProjects As ObservableCollection(Of ProjectInfoModel)
    Public Property FixedProjects As ObservableCollection(Of ProjectInfoModel)
        Get
            Return _FixedProjects
        End Get
        Set(value As ObservableCollection(Of ProjectInfoModel))
            _FixedProjects = value
            AppInfo.FixedProjects = value
            RaisePropertyChanged("FixedProjects")
        End Set
    End Property

    Private Property _ProjectKindList As List(Of String)
    Public Property ProjectKindList As List(Of String)
        Get
            Return Me._ProjectKindList
        End Get
        Set(value As List(Of String))
            Me._ProjectKindList = value
            RaisePropertyChanged("ProjectKindList")
        End Set
    End Property

    Private _ProjectKind As String
    Public Property ProjectKind As String
        Get
            Return _ProjectKind
        End Get
        Set(value As String)
            _ProjectKind = value
            AppInfo.ProjectInfo.Kind = value
            RaisePropertyChanged("ProjectKind")
        End Set
    End Property

    Private _SelectedProject As ProjectInfoModel
    Public Property SelectedProject As ProjectInfoModel
        Get
            Return _SelectedProject
        End Get
        Set(value As ProjectInfoModel)
            _SelectedProject = value
            RaisePropertyChanged("SelectedProject")
        End Set
    End Property

    Private _Message As String
    Public Property Message As String
        Get
            Return _Message
        End Get
        Set(value As String)
            _Message = value
            RaisePropertyChanged("Message")
        End Set
    End Property

    Private Sub _UpdateProjectDirectoryName()
        ProjectDirectoryName = UserDirectoryName & "\" & ProjectName
    End Sub

    'Private Sub _FixedProjectsUpdate(ByVal sender As Object, ByVal e As EventArgs)
    '    Me.FixedProjects = sender
    'End Sub


    ' コマンドプロパティ（Ｐｒｏｊｅｃｔを選択）
    '---------------------------------------------------------------------------------------------'
    Private _SelectProjectCommand As ICommand
    Public ReadOnly Property SelectProjectCommand As ICommand
        Get
            If Me._SelectProjectCommand Is Nothing Then
                Me._SelectProjectCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _SelectProjectCommandExecute,
                    .CanExecuteHandler = AddressOf _SelectProjectCommandCanExecute
                }
                Return Me._SelectProjectCommand
            Else
                Return Me._SelectProjectCommand
            End If
        End Get
    End Property

    Private Sub _CheckSelectProjectCommandEnabled()
        Dim b As Boolean : b = True
        Me._SelectProjectCommandEnableFlag = b
    End Sub

    Private __SelectProjectCommandEnableFlag As Boolean
    Public Property _SelectProjectCommandEnableFlag As Boolean
        Get
            Return Me.__SelectProjectCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__SelectProjectCommandEnableFlag = value
            RaisePropertyChanged("_SelectProjectCommandEnableFlag")
            CType(SelectProjectCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _SelectProjectCommandExecute(ByVal parameter As Object)
        Dim project As ProjectInfoModel
        Dim i = -1
        project = Me.SelectedProject
        Call Me._CheckProject(project)
    End Sub

    Private Function _SelectProjectCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._SelectProjectCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'


    ' コマンドプロパティ（Ｐｒｏｊｅｃｔを開く）
    '---------------------------------------------------------------------------------------------'
    Private _OpenProjectCommand As ICommand
    Public ReadOnly Property OpenProjectCommand As ICommand
        Get
            If Me._OpenProjectCommand Is Nothing Then
                Me._OpenProjectCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _OpenProjectCommandExecute,
                    .CanExecuteHandler = AddressOf _OpenProjectCommandCanExecute
                }
                Return Me._OpenProjectCommand
            Else
                Return Me._OpenProjectCommand
            End If
        End Get
    End Property

    Private Sub _CheckOpenProjectCommandEnabled()
        Dim b As Boolean : b = True
        Me._OpenProjectCommandEnableFlag = b
    End Sub

    Private __OpenProjectCommandEnableFlag As Boolean
    Public Property _OpenProjectCommandEnableFlag As Boolean
        Get
            Return Me.__OpenProjectCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__OpenProjectCommandEnableFlag = value
            RaisePropertyChanged("_OpenProjectCommandEnableFlag")
            CType(OpenProjectCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property


    Private Sub _OpenProjectCommandExecute(ByVal parameter As Object)
        Dim project As ProjectInfoModel
        Dim fbd As New FolderBrowserDialog With {
             .Description = "プロジェクトを開く",
             .RootFolder = Environment.SpecialFolder.Desktop,
             .SelectedPath = Environment.SpecialFolder.Desktop,
             .ShowNewFolderButton = False
        }

        If fbd.ShowDialog() = DialogResult.OK Then
            project = New ProjectInfoModel With {
                .DirectoryName = fbd.SelectedPath
            }
            Call _CheckProject(project)
        End If
    End Sub


    Private Function _OpenProjectCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._OpenProjectCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'


    ' コマンドプロパティ（ＩｎｐｕｔＢｏｘ）
    Private _InputBoxCommand As ICommand
    Public ReadOnly Property InputBoxCommand As ICommand
        Get
            If Me._InputBoxCommand Is Nothing Then
                Me._InputBoxCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _InputBoxCommandExecute,
                    .CanExecuteHandler = AddressOf _InputBoxCommandCanExecute
                }
                Return Me._InputBoxCommand
            Else
                Return Me._InputBoxCommand
            End If
        End Get
    End Property


    'コマンド実行可否のチェック（ＩｎｐｕｔＢｏｘ）
    Private Sub _CheckInputBoxCommandEnabled()
        Dim b As Boolean : b = True
        Me._InputBoxCommandEnableFlag = b
    End Sub


    'コマンド実行可否のフラグ（ＩｎｐｕｔＢｏｘ）
    Private __InputBoxCommandEnableFlag As Boolean
    Public Property _InputBoxCommandEnableFlag As Boolean
        Get
            Return Me.__InputBoxCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__InputBoxCommandEnableFlag = value
            RaisePropertyChanged("_InputBoxCommandEnableFlag")
            CType(InputBoxCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property


    ' コマンド実行（ＩｎｐｕｔＢｏｘ）
    Private Sub _InputBoxCommandExecute(ByVal parameter As Object)
        Dim fbd As New FolderBrowserDialog With {
             .Description = "ディレクトリを指定してください",
             .RootFolder = Environment.SpecialFolder.Desktop,
             .SelectedPath = Environment.SpecialFolder.Desktop,
             .ShowNewFolderButton = True
        }
        If fbd.ShowDialog() = DialogResult.OK Then
            UserDirectoryName = fbd.SelectedPath
        End If
    End Sub


    ' コマンド有効／無効（ＩｎｐｕｔＢｏｘ）
    Private Function _InputBoxCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._InputBoxCommandEnableFlag
    End Function


    ' コマンドプロパティ（Ｐｒｏｊｅｃｔ追加）
    Private _AddProjectCommand As ICommand
    Public ReadOnly Property AddProjectCommand As ICommand
        Get
            If Me._AddProjectCommand Is Nothing Then
                Me._AddProjectCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _AddProjectCommandExecute,
                    .CanExecuteHandler = AddressOf _AddProjectCommandCanExecute
                }
                Return Me._AddProjectCommand
            Else
                Return Me._AddProjectCommand
            End If
        End Get
    End Property


    'コマンド実行可否のチェック（Ｐｒｏｊｅｃｔ追加）
    Private Sub _CheckAddProjectCommandEnabled()
        Dim b As Boolean : b = True

        Me.Message = vbNullString
        If String.IsNullOrEmpty(Me.UserDirectoryName) Then
            b = False
        End If
        If String.IsNullOrEmpty(Me.ProjectName) Then
            b = False
        End If
        If b Then
            If Directory.Exists(Me.ProjectDirectoryName) Then
                b = False
                Me.Message = _DIRECTORY_ALREADY_EXIST
            End If
        End If
        Me._AddProjectCommandEnableFlag = b
    End Sub


    'コマンド実行可否のフラグ（Ｐｒｏｊｅｃｔ追加）
    Private __AddProjectCommandEnableFlag As Boolean
    Public Property _AddProjectCommandEnableFlag As Boolean
        Get
            Return Me.__AddProjectCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__AddProjectCommandEnableFlag = value
            RaisePropertyChanged("_AddProjectCommandEnableFlag")
            CType(AddProjectCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property


    ' コマンド実行（Ｐｒｏｊｅｃｔ追加）
    ' Ａｄｄした時点で、一度セーブを行う
    Private Sub _AddProjectCommandExecute(ByVal parameter As Object)
        Dim i = -1
        Dim jh As New JsonHandler(Of Object)

        Dim project = AppInfo.ProjectInfo

        Dim launcher As New DelegateAction With {
            .CanExecuteHandler = AddressOf project.CheckProjectNotExist,
            .ExecuteHandler = AddressOf project.Launch
        }

        i = launcher.ExecuteIfCan(Nothing)
        If i = 0 Then
            AppInfo.ProjectInfo = project
            Call ProjectSetup()
            Call ModelSetup()
            Call PushProject()
            Call ViewModelSetup()
            Call AllSave()
        Else
            Throw New Exception("Error AddProjectCommandExecute")
        End If
    End Sub


    ' コマンド有効／無効（Ｐｒｏｊｅｃｔ追加）
    Private Function _AddProjectCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._AddProjectCommandEnableFlag
    End Function



    ' プロジェクトをチェックし、正当な場合、プロジェクトをスタートします
    ' Model.Setup() 不要、Model.Initializeする
    Private Sub _CheckProject(ByVal project As ProjectInfoModel)
        Dim msg As String : msg = vbNullString
        Dim jh As New JsonHandler(Of Object)
        Dim vm As ViewModel

        Dim i = project.CheckProject()
        ' 0 ... チェック全てＯＫ
        ' 1 ... ディレクトリが不正
        ' 2 ... プロジェクト情報が不正
        ' 3 ... モデルが不正
        Select Case i
            Case 0
                Call AllLoad(project)
                msg = vbNullString
            Case 1
                Call _DeleteProject(project)
                Call _RemoveFixedProject(project)
                msg = _ISNOT_PROJECT_DIRECTORY & " " & project.DirectoryName
            Case 2
                msg = _PROJECT_LOAD_FAILED & " " & project.DirectoryName
            Case 3
                msg = _PROJECT_LOAD_FAILED & " " & project.DirectoryName
        End Select

        Me.Message = msg
    End Sub


    Private Sub _DeleteProject(ByVal project As ProjectInfoModel)
        For Each p In Me.CurrentProjects
            If p.DirectoryName = project.DirectoryName Then
                Me.CurrentProjects.Remove(p)
                Exit For
            End If
        Next
        Call AppSave()
    End Sub


    Private Sub _ViewInitializing()
        Dim v As ViewItemModel
        Me.UserDirectoryName = AppInfo.ProjectInfo.DirectoryName
        Me.ProjectName = AppInfo.ProjectInfo.Name
        Me.ProjectKindList = AppDirectoryModel.ProjectKindList
        Me.CurrentProjects = AppInfo.CurrentProjects
        Me.FixedProjects = AppInfo.FixedProjects

        Dim fvm = New FrameViewModel With {
            .MainViewContent = New ViewItemModel With {
                .Name = "プロジェクト選択",
                .Content = Me
            }
        }
        Call ViewModel.VisualizeView(fvm)
    End Sub

    Private Sub _RemoveFixedProject(ByVal project As ProjectInfoModel)
        For Each p In Me.FixedProjects
            If p.DirectoryName = project.DirectoryName Then
                Me.FixedProjects.Remove(p)
                Exit For
            End If
        Next

        Call AppSave()
    End Sub

    Private Sub _PushFixedProject(ByVal project As ProjectInfoModel)
        Dim [old] As ProjectInfoModel()
        ReDim old(Me.FixedProjects.Count - 1)

        Me.FixedProjects.CopyTo([old], 0)
        Do Until True = False
            For Each p In Me.FixedProjects
                Me.FixedProjects.Remove(p)
                Exit For
            Next
            If Me.FixedProjects.Count = 0 Then
                Exit Do
            End If
        Loop
        Me.FixedProjects.Add(project)
        If project.[FixedIndex] = 0 Then
            Call _AssignFixedProjectIndex(project)
        End If

        For Each p In [old]
            If project.[FixedIndex] <> p.[FixedIndex] Then
                Me.FixedProjects.Add(p)
            End If
            If Me.FixedProjects.Count >= 5 Then
                Exit For
            End If
        Next
        Call AppSave()
    End Sub

    Private Sub _AssignFixedProjectIndex(ByRef project As ProjectInfoModel)
        Dim idx = 1
        Dim b = False

        Do Until True = False
            b = False
            For Each p In Me.FixedProjects
                If p.[Index] = idx Then
                    idx += 1
                    b = True
                    Exit For
                End If
            Next
            If Not b Then
                Exit Do
            End If
        Loop
        project.[FixedIndex] = idx
    End Sub

    '--- プロジェクトのピン解除 ------------------------------------------------------------------'
    Private Sub _RemoveFixedProjectRequestedReview(ByVal p As ProjectInfoModel, ByVal e As System.EventArgs)
        Call _RemoveFixedProjectRequestAccept(p)
    End Sub

    Private Sub _RemoveFixedProjectRequestAccept(ByVal p As ProjectInfoModel)
        Call _RemoveFixedProject(p)
    End Sub

    Private Sub _RemoveFixedProjectAddHandler()
        AddHandler _
            DelegateEventListener.Instance.RemoveFixedProjectRequested,
            AddressOf Me._RemoveFixedProjectRequestedReview
    End Sub
    '---------------------------------------------------------------------------------------------'

    '--- プロジェクトのピン止め ------------------------------------------------------------------'
    Private Sub _FixProjectRequestedReview(ByVal p As ProjectInfoModel, ByVal e As System.EventArgs)
        Call _FixProjectRequestAccept(p)
    End Sub

    Private Sub _FixProjectRequestAccept(ByVal p As ProjectInfoModel)
        Call _PushFixedProject(p)
    End Sub

    Private Sub _FixProjectAddHandler()
        AddHandler _
            DelegateEventListener.Instance.FixProjectRequested,
            AddressOf Me._FixProjectRequestedReview
    End Sub
    '---------------------------------------------------------------------------------------------'

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel,
                                    ByRef vm As ViewModel)
        InitializeHandler _
            = AddressOf _ViewInitializing
        CheckCommandEnabledHandler _
            = [Delegate].Combine(
                New Action(AddressOf _CheckAddProjectCommandEnabled),
                New Action(AddressOf _CheckInputBoxCommandEnabled),
                New Action(AddressOf _CheckOpenProjectCommandEnabled),
                New Action(AddressOf _CheckSelectProjectCommandEnabled)
            )

        [AddHandler] = [Delegate].Combine(
            New Action(AddressOf _FixProjectAddHandler),
            New Action(AddressOf _RemoveFixedProjectAddHandler)
        )

        Call BaseInitialize(app, vm)
    End Sub
End Class
