﻿Imports System.Collections.ObjectModel
Imports System.Windows.Forms
Imports System.IO

Public Class MenuViewModel
    Inherits BaseViewModel2

    ' コマンドプロパティ（Ｐｒｏｊｅｃｔ設定画面表示）
    '---------------------------------------------------------------------------------------------'
    Private _ShowProjectSettingCommand As ICommand
    Public ReadOnly Property ShowProjectSettingCommand As ICommand
        Get
            If Me._ShowProjectSettingCommand Is Nothing Then
                Me._ShowProjectSettingCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ShowProjectSettingCommandExecute,
                    .CanExecuteHandler = AddressOf _ShowProjectSettingCommandCanExecute
                }
                Return Me._ShowProjectSettingCommand
            Else
                Return Me._ShowProjectSettingCommand
            End If
        End Get
    End Property

    Private Sub _CheckShowProjectSettingCommandEnabled()
        Dim b As Boolean : b = True
        Me._ShowProjectSettingCommandEnableFlag = b
    End Sub

    Private __ShowProjectSettingCommandEnableFlag As Boolean
    Public Property _ShowProjectSettingCommandEnableFlag As Boolean
        Get
            Return Me.__ShowProjectSettingCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ShowProjectSettingCommandEnableFlag = value
            RaisePropertyChanged("_ShowProjectSettingCommandEnableFlag")
            CType(ShowProjectSettingCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _ShowProjectSettingCommandExecute(ByVal parameter As Object)
        Dim v As ViewItemModel
        Dim pvm = New ProjectViewModel
        pvm.Initialize(AppInfo, ViewModel)
        v = New ViewItemModel With {
            .Content = pvm,
            .Name = pvm.GetType.Name,
            .FrameType = MultiViewModel.MAIN_FRAME,
            .LayoutType = ViewModel.MULTI_VIEW,
            .ViewType = MultiViewModel.TAB_VIEW,
            .OpenState = True
        }
        Call AddViewItem(v)
        Call AddView(v)
    End Sub

    Private Function _ShowProjectSettingCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ShowProjectSettingCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

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
        Call AllSave()
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
                .Kind = AppInfo.ProjectInfo.Kind
            }
            Call project.Launch()
            Call AllSave(project)
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
                Call InitializeViewContent()
                Call ProjectLoad(project)
                Call PushProject(AppInfo.ProjectInfo)
                Call ProjectModelLoad()
                Call ModelSetup()
                Call ProjectViewModelLoad()
                Call ViewModelSetup()
                Call AppSave()
            Case 1
                Call AppSave()
            Case 2
            Case 3
        End Select
    End Sub

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel,
                                    ByRef vm As ViewModel)
        CheckCommandEnabledHandler = [Delegate].Combine(
            New Action(AddressOf _CheckOpenProjectCommandEnabled),
            New Action(AddressOf _CheckSaveProjectCommandEnabled),
            New Action(AddressOf _CheckResaveProjectCommandEnabled),
            New Action(AddressOf _CheckShowProjectSettingCommandEnabled)
        )

        Call BaseInitialize(app, vm)
    End Sub
End Class
