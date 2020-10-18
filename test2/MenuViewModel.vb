Imports System.Collections.ObjectModel
Imports System.Windows.Forms

Public Class MenuViewModel
    Inherits BaseViewModel2

    Public Overrides ReadOnly Property ViewType As String
        Get
            Return ViewModel.MENU_VIEW
        End Get
    End Property


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
                ProjectInfo = jh.ModelLoad(Of ProjectInfoModel)(project.ProjectInfoFileName)
                Model = jh.ModelLoad(Of Model)(project.ModelFileName)
                Call Model.Initialize(ProjectInfo)
                Call ViewModel.Setup(Model, ViewModel, AppInfo, ProjectInfo)
            Case 1
                'Call _DeleteProject(project)
                Call AppInfo.AppSave()
            Case 2
            Case 3
        End Select
    End Sub

    Public Sub Initialize(ByRef m As Model,
                          ByRef vm As ViewModel,
                          ByRef adm As AppDirectoryModel,
                          ByRef pim As ProjectInfoModel)

        'InitializeHandler = AddressOf _ViewInitialize
        CheckCommandEnabledHandler = [Delegate].Combine(
            New Action(AddressOf _CheckOpenProjectCommandEnabled)
        )

        Call BaseInitialize(m, vm, adm, pim)
    End Sub
End Class
