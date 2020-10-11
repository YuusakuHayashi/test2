﻿Imports System.Windows.Forms
Imports System.Collections.ObjectModel
Imports System.IO

Public Class UserDirectoryViewModel
    Inherits BaseViewModel2(Of Object)

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
            ProjectInfo.Name = value
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
            ProjectInfo.DirectoryName = value
            RaisePropertyChanged("ProjectDirectoryName")
        End Set
    End Property

    ' 
    Private _CurrentProjects As ObservableCollection(Of ProjectInfoModel)
    Public Property CurrentProjects As ObservableCollection(Of ProjectInfoModel)
        Get
            Return _CurrentProjects
        End Get
        Set(value As ObservableCollection(Of ProjectInfoModel))
            _CurrentProjects = value
            AppInfo.CurrentProjects = value
            RaisePropertyChanged("CurrentProjects")
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
            ProjectInfo.Kind = value
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
        project = Me.SelectedProject
        Call _CheckProject(project)
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
    Private Sub _AddProjectCommandExecute(ByVal parameter As Object)
        Dim i = -1
        Dim rtn As Action(Of ProjectInfoModel)

        ' 新規プロジェクト情報
        Dim project As New ProjectInfoModel With {
            .DirectoryName = Me.ProjectDirectoryName,
            .Name = Me.ProjectName,
            .Kind = Me.ProjectKind
        }

        Dim launcher As New DelegateAction With {
            .CanExecuteHandler = AddressOf _CheckProjectNotExist,
            .ExecuteHandler = AddressOf _LaunchProject
        }

        i = launcher.ExecuteIfCan(project)
        If i = 0 Then
            AppInfo.CurrentProjects = StackModule.Push(Of ObservableCollection(Of ProjectInfoModel), ProjectInfoModel)(project, AppInfo.CurrentProjects, 5)
            Call AppInfo.ModelSave(
                AppDirectoryModel.ModelFileName,
                AppInfo
            )
            ProjectInfo = project
            Call ProjectInfo.ModelSave(
                project.ProjectInfoFileName,
                ProjectInfo
            )
            Call Model.InitializeModels(project.Kind)
            Call Model.ModelSave(
                project.ModelFileName,
                Model
            )
            Call ViewModel.InitializeViewModelsOfProject(
                project.Kind,
                Model,
                ViewModel,
                AppInfo,
                ProjectInfo
            )
            Me.CurrentProjects = AppInfo.CurrentProjects
        End If
    End Sub


    ' コマンド有効／無効（Ｐｒｏｊｅｃｔ追加）
    Private Function _AddProjectCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._AddProjectCommandEnableFlag
    End Function


    ' プロジェクトをチェックし、正当な場合、プロジェクトをスタートします
    Private Sub _CheckProject(ByVal project As ProjectInfoModel)
        Dim msg As String : msg = vbNullString
        Dim i As Integer : i = -1
        Dim i2 As Integer : i = -1
        Dim projectInfoLoader As New DelegateAction With {
            .CanExecuteHandler = AddressOf _CheckProjectExist,
            .CanExecuteHandler2 = AddressOf _CheckProjectInfo,
            .ExecuteHandler = AddressOf _LoadProjectInfo
        }

        Dim modelLoader As New DelegateAction With {
            .CanExecuteHandler = AddressOf _CheckModel,
            .ExecuteHandler = AddressOf _LoadModel
        }

        Dim projectDeleter As New DelegateAction With {
            .ExecuteHandler = AddressOf _DeleteProject
        }

        i = projectInfoLoader.ExecuteIfCan(project)
        If i = 0 Then
            i2 = modelLoader.ExecuteIfCan(project)
            If i2 = 0 Then
                Call ViewModel.InitializeViewModelsOfProject(
                    ProjectInfo.Kind,
                    Model,
                    ViewModel,
                    AppInfo,
                    ProjectInfo
                )
                msg = vbNullString
            Else
                msg = _PROJECT_LOAD_FAILED & " " & project.DirectoryName
            End If
        Else
            Call _DeleteProject(project)
            Call AppInfo.ModelSave(AppDirectoryModel.ModelFileName, AppInfo)
            msg = _ISNOT_PROJECT_DIRECTORY & " " & project.DirectoryName
        End If

        Me.Message = msg
    End Sub

    Private Sub _LaunchProject(ByVal project As ProjectInfoModel)
        Call project.Launch()
    End Sub

    Private Overloads Function _CheckProjectExist(ByVal project As ProjectInfoModel) As Boolean
        Dim i = -1
        Dim b = False

        i = project.CheckStructure()
        If i = 0 Then
            b = True
        End If

        _CheckProjectExist = b
    End Function

    Private Overloads Function _CheckProjectNotExist(ByVal project As ProjectInfoModel) As Boolean
        Dim i = -1
        Dim b = False

        i = project.CheckStructure()
        If i = 1000 Then
            b = True
        End If

        _CheckProjectNotExist = b
    End Function

    Private Function _CheckProjectInfo(ByVal project As ProjectInfoModel) As Boolean
        Dim jh As New JsonHandler(Of Nullable)
        Return (jh.CheckModel(Of ProjectInfoModel)(project.ProjectInfoFileName))
    End Function

    Private Function _CheckModel(ByVal project As ProjectInfoModel) As Boolean
        Dim jh As New JsonHandler(Of Nullable)
        Return (jh.CheckModel(Of Model)(project.ModelFileName))
    End Function

    Private Sub _SaveAppInfo(ByVal app As AppDirectoryModel)
    End Sub

    Private Sub _LoadProjectInfo(ByVal project As ProjectInfoModel)
        Dim jh As New JsonHandler(Of Nullable)
        Me.ProjectInfo = jh.ModelLoad(Of ProjectInfoModel)(project.ProjectInfoFileName)
    End Sub

    Private Sub _LoadModel(ByVal project As ProjectInfoModel)
        Dim jh As New JsonHandler(Of Nullable)
        Me.Model = jh.ModelLoad(Of Model)(project.ModelFileName)
    End Sub

    Private Sub _DeleteProject(ByVal project As ProjectInfoModel)
        Dim pim As ProjectInfoModel
        For Each p In Me.CurrentProjects
            If p.DirectoryName = project.DirectoryName Then
                pim = p
                Exit For
            End If
        Next
        If pim IsNot Nothing Then
            Me.CurrentProjects.Remove(pim)
        End If
    End Sub

    Protected Overrides Sub ViewInitializing()
        Me.UserDirectoryName = ProjectInfo.DirectoryName
        Me.ProjectName = ProjectInfo.Name
        Me.ProjectKindList = AppDirectoryModel.ProjectKindList
        Me.CurrentProjects = AppInfo.CurrentProjects
        If Me.CurrentProjects Is Nothing Then
            Me.CurrentProjects = New ObservableCollection(Of ProjectInfoModel)
        End If
        Me.ViewModel.SetContent(ViewModel.MAIN_VIEW, Me.GetType.Name, New NormalViewModel With {.Content = Me})
    End Sub


    Public Sub MyInitializing(ByRef m As Model,
                              ByRef vm As ViewModel,
                              ByRef adm As AppDirectoryModel,
                              ByRef pim As ProjectInfoModel)
        Dim ip As InitializingProxy
        ip = AddressOf ViewInitializing

        Dim ccep(2) As CheckCommandEnabledProxy
        Dim ccep2 As CheckCommandEnabledProxy
        ccep(0) = AddressOf Me._CheckInputBoxCommandEnabled
        ccep(1) = AddressOf Me._CheckOpenProjectCommandEnabled
        ccep(2) = AddressOf Me._CheckSelectProjectCommandEnabled
        ccep2 = [Delegate].Combine(ccep)

        Call Initializing(m, vm, adm, pim, ip, ccep2)
    End Sub

    ' ビューモデルは必ず、Initializingメソッドを呼び出すこと
    Sub New()
    End Sub
End Class