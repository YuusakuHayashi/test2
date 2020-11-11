Imports System.Collections.ObjectModel
Imports System.Windows.Forms
Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class MenuViewModel
    Inherits BaseViewModel2

    ' コマンドプロパティ（プロジェクトビュー表示）
    '---------------------------------------------------------------------------------------------'
    Private _ShowProjectExplorerCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property ShowProjectExplorerCommand As ICommand
        Get
            If Me._ShowProjectExplorerCommand Is Nothing Then
                Me._ShowProjectExplorerCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ShowProjectExplorerCommandExecute,
                    .CanExecuteHandler = AddressOf _ShowProjectExplorerCommandCanExecute
                }
                Return Me._ShowProjectExplorerCommand
            Else
                Return Me._ShowProjectExplorerCommand
            End If
        End Get
    End Property

    Private Sub _CheckShowProjectExplorerCommandEnabled()
        Dim b As Boolean : b = True
        Me._ShowProjectExplorerCommandEnableFlag = b
    End Sub

    Private __ShowProjectExplorerCommandEnableFlag As Boolean
    Private Property _ShowProjectExplorerCommandEnableFlag As Boolean
        Get
            Return Me.__ShowProjectExplorerCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ShowProjectExplorerCommandEnableFlag = value
            RaisePropertyChanged("_ShowProjectExplorerCommandEnableFlag")
            CType(ShowProjectExplorerCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _ShowProjectExplorerCommandExecute(ByVal parameter As Object)
        Call _ShowProjectExplorer(ViewModel.Content)
    End Sub

    Private Function _ShowProjectExplorerCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ShowProjectExplorerCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Sub _ShowProjectExplorer(ByRef fvm As FlexibleViewModel)
    End Sub

    ' コマンドプロパティ（プロジェクトビュー表示）
    '---------------------------------------------------------------------------------------------'
    Private _ShowViewExplorerCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property ShowViewExplorerCommand As ICommand
        Get
            If Me._ShowViewExplorerCommand Is Nothing Then
                Me._ShowViewExplorerCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ShowViewExplorerCommandExecute,
                    .CanExecuteHandler = AddressOf _ShowViewExplorerCommandCanExecute
                }
                Return Me._ShowViewExplorerCommand
            Else
                Return Me._ShowViewExplorerCommand
            End If
        End Get
    End Property

    Private Sub _CheckShowViewExplorerCommandEnabled()
        Dim b As Boolean : b = True
        Me._ShowViewExplorerCommandEnableFlag = b
    End Sub

    Private __ShowViewExplorerCommandEnableFlag As Boolean
    Private Property _ShowViewExplorerCommandEnableFlag As Boolean
        Get
            Return Me.__ShowViewExplorerCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ShowViewExplorerCommandEnableFlag = value
            RaisePropertyChanged("_ShowViewExplorerCommandEnableFlag")
            CType(ShowViewExplorerCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _ShowViewExplorerCommandExecute(ByVal parameter As Object)
        Dim names = _SearchViewExplorer(New List(Of String), ViewModel.Content)
        If names.Count > 0 Then
            Call _ShowViewExplorer(ViewModel.Content, names)
        End If
    End Sub

    Private Function _ShowViewExplorerCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ShowViewExplorerCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Overloads Function _SearchViewExplorer(ByVal NewNames As List(Of String),
                                                   ByVal fvm As FlexibleViewModel) As List(Of String)
        Dim oldcount = NewNames.Count
        If NewNames.Count = oldcount Then
            If fvm.MainViewContent IsNot Nothing Then
                NewNames = _SearchViewExplorer(NewNames, fvm.MainViewContent)
            End If
        End If
        If NewNames.Count = oldcount Then
            If fvm.RightViewContent IsNot Nothing Then
                NewNames = _SearchViewExplorer(NewNames, fvm.RightViewContent)
            End If
        End If
        If NewNames.Count = oldcount Then
            If fvm.BottomViewContent IsNot Nothing Then
                NewNames = _SearchViewExplorer(NewNames, fvm.BottomViewContent)
            End If
        End If
        Return NewNames
    End Function

    Private Overloads Function _SearchViewExplorer(ByVal NewNames As List(Of String),
                                                   ByVal vim As ViewItemModel) As List(Of String)
        Dim fvm As FlexibleViewModel
        Dim tvm As TabViewModel

        Dim OldNames As New List(Of String)
        '何故か和が求まらない
        'OldNames.Concat(NewNames)
        For Each nn In NewNames
            OldNames.Add(nn)
        Next

        NewNames.Add(vim.Name)

        Select Case vim.ModelName
            Case "FlexibleViewModel"
                fvm = CType(vim.Content, FlexibleViewModel)
                NewNames = _SearchViewExplorer(NewNames, fvm)
            Case "TabViewModel"
                tvm = CType(vim.Content, TabViewModel)
                NewNames = _SearchViewExplorer(NewNames, tvm)
            Case "ViewExplorerViewModel"
                'NewNames.Add(vim.Name)
            Case Else
                NewNames = OldNames
        End Select
        _SearchViewExplorer = NewNames
    End Function

    Private Overloads Function _SearchViewExplorer(ByVal NewNames As List(Of String),
                                                   ByVal tvm As TabViewModel) As List(Of String)
        Dim oldcount = NewNames.Count
        For Each vt In tvm.ViewContentTabs
            If NewNames.Count = oldcount Then
                NewNames = _SearchViewExplorer(NewNames, vt)
            End If
        Next
        Return NewNames
    End Function

    Private Function _PopData(Of T As {New, IList})(ByVal old As T) As T
        Dim [new] As New T
        For Each o In old
            If old.IndexOf(o) <> 0 Then
                [new].Add(o)
            End If
        Next
        _PopData = [new]
    End Function

    Private Sub _ShowViewExplorerContent(ByRef obj As Object, ByVal vim As ViewItemModel, names As List(Of String))
        Dim fvm As FlexibleViewModel
        Dim tvm As TabViewModel
        If vim.Name = names(0) Then
            If obj Is Nothing Then
                obj = vim.Content
            Else
                names = _PopData(names)
                Select Case vim.ModelName
                    Case "FlexibleViewModel"
                        fvm = CType(obj, FlexibleViewModel)
                        Call _ShowViewExplorerFromFlex(fvm, names)
                    Case "TabViewModel"
                        tvm = CType(obj, TabViewModel)
                        Call _ShowViewExplorer(tvm, names)
                    Case Else
                        Call _ShowViewExplorerContent(obj, names)
                End Select
            End If
        End If
    End Sub

    Private Sub _ShowViewExplorerFromTabs(ByRef tvm As FlexibleViewModel, names As List(Of String))
        Dim idx As Integer
        Dim fvm As FlexibleViewModel
        Dim tvm2 As TabViewModel

        For Each vt In tvm.ViewContentTabs
            idx = -1
            If vt.Name = names(0) Then
                names = _PopData(names)
                Select Case vt.ModelName
                    Case "FlexibleViewModel"
                        fvm = CType(obj, FlexibleViewModel)
                        Call _ShowViewExplorerFromFlex(fvm, names)
                    Case "TabViewModel"
                        tvm2 = CType(obj, TabViewModel)
                        Call _ShowViewExplorerFromTabs(tvm2, names)
                    Case Else
                        Call _ShowViewExplorerContent(obj, names)
                End Select
                Exit For
            End If
        Next
    End If

    Private Sub _ShowViewExplorerFromFlex(ByRef fvm As FlexibleViewModel, names As List(Of String))
        If fvm.MainViewContent IsNot Nothing Then
            If fvm.MainViewContent.Name = names(0) Then
                If fvm.MainContent Is Nothing Then
                    fvm.MainContent = fvm.MainViewContent.Content
                Else
                    names = _PopData(names)
                    Call _ShowViewExplorerContent(fvm.MainContent, fvm.MainViewContent)
                End If
            End If
        End If
        If fvm.RightViewContent IsNot Nothing Then
            If fvm.RightViewContent.Name = names(0) Then
                If fvm.RightContent Is Nothing Then
                    fvm.RightContent = fvm.RightViewContent.Content
                Else
                    names = _PopData(names)
                    Call _ShowViewExplorerContent(fvm.RightContent, fvm.RightViewContent)
                End If
            End If
        End If
        If fvm.BottomViewContent IsNot Nothing Then
            If fvm.BottomViewContent.Name = names(0) Then
                If fvm.BottomContent Is Nothing Then
                    fvm.BottomContent = fvm.BottomViewContent.Content
                Else
                    names = _PopData(names)
                    Call _ShowViewExplorerContent(fvm.BottomContent, fvm.BottomViewContent)
                End If
            End If
        End If
    End Sub


    ' コマンドプロパティ（Ｐｒｏｊｅｃｔ設定画面表示）
    '---------------------------------------------------------------------------------------------'
    Private _ShowProjectSettingCommand As ICommand
    <JsonIgnore>
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
    Private Property _ShowProjectSettingCommandEnableFlag As Boolean
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
        'Dim v As ViewItemModel
        'Dim pvm = New ProjectViewModel
        'pvm.Initialize(AppInfo, ViewModel)
        'v = New ViewItemModel With {
        '    .Content = pvm,
        '    .Name = pvm.GetType.Name,
        '    .ModelName = pvm.GetType.Name,
        '    .FrameType = MultiViewModel.MAIN_FRAME,
        '    .LayoutType = ViewModel.MULTI_VIEW,
        '    .ViewType = MultiViewModel.TAB_VIEW,
        '    .OpenState = True
        '}
        'ViewModel.Views.Add(v)
        'Call AddView(v)
    End Sub

    Private Function _ShowProjectSettingCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ShowProjectSettingCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    ' コマンドプロパティ（Ｐｒｏｊｅｃｔを上書き）
    '---------------------------------------------------------------------------------------------'
    Private _ResaveProjectCommand As ICommand
    <JsonIgnore>
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
    Private Property _ResaveProjectCommandEnableFlag As Boolean
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
    <JsonIgnore>
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
    Private Property _SaveProjectCommandEnableFlag As Boolean
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
    <JsonIgnore>
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
    Private Property _OpenProjectCommandEnableFlag As Boolean
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
                Call AllLoad(project)
                'Call ProjectLoad(project)
                'Call PushProject(AppInfo.ProjectInfo)
                'Call ProjectModelLoad()
                'Call ModelSetup()
                'Call ProjectViewModelLoad()
                'Call ViewModelSetup()
                'Call AppSave()
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
            New Action(AddressOf _CheckShowProjectSettingCommandEnabled),
            New Action(AddressOf _CheckShowProjectExplorerCommandEnabled),
            New Action(AddressOf _CheckShowViewExplorerCommandEnabled)
        )

        Call BaseInitialize(app, vm)
    End Sub
End Class
