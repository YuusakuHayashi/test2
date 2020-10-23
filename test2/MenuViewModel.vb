Imports System.Collections.ObjectModel
Imports System.Windows.Forms
Imports System.IO

Public Class MenuViewModel
    Inherits BaseViewModel2

    Public Overrides ReadOnly Property FrameType As String
        Get
            Return ViewModel.MENU_FRAME
        End Get
    End Property

    ' コマンドプロパティ（Ｐｒｏｊｅｃｔを上書き）
    '---------------------------------------------------------------------------------------------'
    Private _ResaveProjectCommand As ICommand
    Public ReadOnly Property ResaveProjectCommand As ICommand
        Get
            If Me._ResaveProjectCommand Is Nothing Then
                Me._ResaveProjectCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ResaveProjectCommandExecute,
                    .CanExecuteHandler = AddressOf _ResaveProjectCommandCanExecute
                }
                Return Me._ResaveProjectCommand
            Else
                Return Me._ResaveProjectCommand
            End If
        End Get
    End Property

    Private Sub _CheckResaveProjectCommandEnabled()
        Dim b As Boolean : b = True
        Me._ResaveProjectCommandEnableFlag = b
    End Sub

    Private __ResaveProjectCommandEnableFlag As Boolean
    Public Property _ResaveProjectCommandEnableFlag As Boolean
        Get
            Return Me.__ResaveProjectCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ResaveProjectCommandEnableFlag = value
            RaisePropertyChanged("_ResaveProjectCommandEnableFlag")
            CType(ResaveProjectCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _ResaveProjectCommandExecute(ByVal parameter As Object)
        Call Model.ModelSave(ProjectInfo.ModelFileName, Model)
        Call ViewModel.ModelSave(ProjectInfo.ViewModelFileName, ViewModel)
        Call ProjectInfo.ProjectSave()
        Call AppInfo.PushProject(ProjectInfo)
        Call AppInfo.AppSave()
    End Sub

    Private Function _ResaveProjectCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ResaveProjectCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    ' コマンドプロパティ（Ｐｒｏｊｅｃｔを保存）
    '---------------------------------------------------------------------------------------------'
    Private _SaveProjectCommand As ICommand
    Public ReadOnly Property SaveProjectCommand As ICommand
        Get
            If Me._SaveProjectCommand Is Nothing Then
                Me._SaveProjectCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _SaveProjectCommandExecute,
                    .CanExecuteHandler = AddressOf _SaveProjectCommandCanExecute
                }
                Return Me._SaveProjectCommand
            Else
                Return Me._SaveProjectCommand
            End If
        End Get
    End Property

    Private Sub _CheckSaveProjectCommandEnabled()
        Dim b As Boolean : b = True
        Me._SaveProjectCommandEnableFlag = b
    End Sub

    Private __SaveProjectCommandEnableFlag As Boolean
    Public Property _SaveProjectCommandEnableFlag As Boolean
        Get
            Return Me.__SaveProjectCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__SaveProjectCommandEnableFlag = value
            RaisePropertyChanged("_SaveProjectCommandEnableFlag")
            CType(SaveProjectCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _SaveProjectCommandExecute(ByVal parameter As Object)
        Dim i = -1
        Dim [dir] As String
        Dim [root] As String
        Dim jh As New JsonHandler(Of Object)
        Dim project As ProjectInfoModel
        Dim fbd As New FolderBrowserDialog With {
             .Description = "プロジェクトを保存",
             .RootFolder = Environment.SpecialFolder.Desktop,
             .SelectedPath = Environment.SpecialFolder.Desktop,
             .ShowNewFolderButton = True
        }

        If fbd.ShowDialog() = DialogResult.OK Then
            [dir] = fbd.SelectedPath
            [root] = [dir].Replace((Path.GetDirectoryName(dir) & "\"), vbNullString)
            project = New ProjectInfoModel With {
                .DirectoryName = [dir],
                .Name = [root],
                .Kind = ProjectInfo.Kind
            }
            Call project.Launch()
            Call jh.ModelSave(Of ProjectInfoModel)(project.ProjectInfoFileName, project)
            Call jh.ModelSave(Of Model)(project.ModelFileName, Model)
            AppInfo.CurrentProjects = StackModule.Push(Of ObservableCollection(Of ProjectInfoModel), ProjectInfoModel)(project, AppInfo.CurrentProjects, 5)
            Call AppInfo.AppSave()
        End If
    End Sub


    Private Function _SaveProjectCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._SaveProjectCommandEnableFlag
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
        Dim i = -1
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

    ' UserDirectoryModelの同名サブルーチンをフォーク
    Private Sub _CheckProject(ByVal project As ProjectInfoModel)
        Dim msg As String : msg = vbNullString
        Dim jh As New JsonHandler(Of Object)

        Dim i = project.CheckProject()
        ' 0 ... チェック全てＯＫ
        ' 1 ... ディレクトリが不正
        ' 2 ... プロジェクト情報が不正
        ' 3 ... モデルが不正
        Select Case i
            Case 0
                ProjectInfo = project.ModelLoad(project.ProjectInfoFileName)
                AppInfo.PushProject(ProjectInfo)
                Model = jh.ModelLoad(Of Model)(ProjectInfo.ModelFileName)
                Call Model.Initialize(ProjectInfo)
                Call Setup(Model, ViewModel, AppInfo, ProjectInfo)
            Case 1
                'Call _DeleteProject(project)
                Call AppInfo.AppSave()
            Case 2
            Case 3
        End Select
    End Sub


    Public Overrides Sub Initialize(ByRef m As Model, ByRef vm As ViewModel, ByRef adm As AppDirectoryModel, ByRef pim As ProjectInfoModel)
        CheckCommandEnabledHandler = [Delegate].Combine(
            New Action(AddressOf _CheckOpenProjectCommandEnabled),
            New Action(AddressOf _CheckSaveProjectCommandEnabled),
            New Action(AddressOf _CheckResaveProjectCommandEnabled)
        )

        Call BaseInitialize(m, vm, adm, pim)
    End Sub
End Class
