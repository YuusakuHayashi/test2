Imports System.Windows.Forms
Imports System.Collections.ObjectModel

Public Class UserDirectoryViewModel : Inherits BaseViewModel2

    Private _UserDirectoryName As String
    Public Property UserDirectoryName As String
        Get
            Return _UserDirectoryName
        End Get
        Set(value As String)
            _UserDirectoryName = value
            RaisePropertyChanged("UserDirectoryName")
            Call _UpdateProjectDirectoryName()
            Call _CheckProjectAddCommandEnabled()
        End Set
    End Property

    Private _ProjectName As String
    Public Property ProjectName As String
        Get
            Return _ProjectName
        End Get
        Set(value As String)
            _ProjectName = value
            RaisePropertyChanged("ProjectName")
            Call _UpdateProjectDirectoryName()
            Call _CheckProjectAddCommandEnabled()
        End Set
    End Property

    Private _ProjectDirectoryName As String
    Public Property ProjectDirectoryName As String
        Get
            Return _ProjectDirectoryName
        End Get
        Set(value As String)
            _ProjectDirectoryName = value
            RaisePropertyChanged("ProjectDirectoryName")
        End Set
    End Property

    Private Sub _UpdateProjectDirectoryName()
        ProjectDirectoryName = UserDirectoryName & "\" & ProjectName
    End Sub

    Private _CurrentProjects As ObservableCollection(Of AppDirectoryModel.ProjectModel)
    Public Property CurrentProjects As ObservableCollection(Of AppDirectoryModel.ProjectModel)
        Get
            Return _CurrentProjects
        End Get
        Set(value As ObservableCollection(Of ProjectInfoModel.ProjectModel))
            _CurrentProjects = value
            RaisePropertyChanged("CurrentProjects")
        End Set
    End Property

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
    Private _ProjectAddCommand As ICommand
    Public ReadOnly Property ProjectAddCommand As ICommand
        Get
            If Me._ProjectAddCommand Is Nothing Then
                Me._ProjectAddCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ProjectAddCommandExecute,
                    .CanExecuteHandler = AddressOf _ProjectAddCommandCanExecute
                }
                Return Me._ProjectAddCommand
            Else
                Return Me._ProjectAddCommand
            End If
        End Get
    End Property


    'コマンド実行可否のチェック（Ｐｒｏｊｅｃｔ追加）
    Private Sub _CheckProjectAddCommandEnabled()
        Dim b As Boolean : b = True
        If String.IsNullOrEmpty(Me.UserDirectoryName) Then
            b = False
        End If
        If String.IsNullOrEmpty(Me.ProjectName) Then
            b = False
        End If
        Me._ProjectAddCommandEnableFlag = b
    End Sub


    'コマンド実行可否のフラグ（Ｐｒｏｊｅｃｔ追加）
    Private __ProjectAddCommandEnableFlag As Boolean
    Public Property _ProjectAddCommandEnableFlag As Boolean
        Get
            Return Me.__ProjectAddCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ProjectAddCommandEnableFlag = value
            RaisePropertyChanged("_ProjectAddCommandEnableFlag")
            CType(ProjectAddCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property


    ' コマンド実行（Ｐｒｏｊｅｃｔ追加）
    Private Sub _ProjectAddCommandExecute(ByVal parameter As Object)
        Dim [new] As ObservableCollection(Of AppDirectoryModel.ProjectModel)
        Dim elm As New AppDirectoryModel.ProjectModel With {
            .Directory = Me.ProjectDirectoryName,
            .Name = ProjectName
        }
        [new] = StackModule.Push(Of ObservableCollection(Of AppDirectoryModel.ProjectModel), AppDirectoryModel.ProjectModel)(elm, Me.CurrentProjects)

        Me.CurrentProjects = [new]
    End Sub


    ' コマンド有効／無効（Ｐｒｏｊｅｃｔ追加）
    Private Function _ProjectAddCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ProjectAddCommandEnableFlag
    End Function

    Protected Overrides Sub ViewInitializing()
        Me.UserDirectoryName = Me.ProjectInfo.UserDirectoryName
        Me.ProjectName = Me.ProjectInfo.ProjectName
        Me.CurrentProjects = Me.ProjectInfo.CurrentProjects
        If Me.CurrentProjects Is Nothing Then
            Me.CurrentProjects = New ObservableCollection(Of AppDirectoryModel.ProjectModel)
        End If
        'Me.CurrentProjects = New ObservableCollection(Of AppDirectoryModel.ProjectModel) From {
        '    New AppDirectoryModel.ProjectModel With {.Name = "Test1"},
        '    New AppDirectoryModel.ProjectModel With {.Name = "Test2"},
        '    New AppDirectoryModel.ProjectModel With {.Name = "Test3"}
        '}
    End Sub

    ' ビューモデルへの自身の登録を行うメソッドを慣習的にこの名前でオーバライドしています
    Protected Overrides Sub ContextModelCheck()
        ViewModel.SetContext(ViewModel.MAIN_VIEW, Me.GetType.Name, Me)
    End Sub

    ' ビューモデルは必ず、Initializingメソッドを呼び出すこと
    Sub New(ByRef m As Model2,
            ByRef vm As ViewModel,
            ByRef pim As ProjectInfoModel)

        Dim ip As InitializingProxy
        ip = AddressOf ViewInitializing

        Dim cmcp(0) As ContextModelCheckProxy
        Dim cmcp2 As ContextModelCheckProxy
        cmcp(0) = AddressOf Me.ContextModelCheck
        cmcp2 = [Delegate].Combine(cmcp)

        Dim ccep(0) As CheckCommandEnabledProxy
        Dim ccep2 As CheckCommandEnabledProxy
        ccep(0) = AddressOf Me._CheckInputBoxCommandEnabled
        ccep2 = [Delegate].Combine(ccep)

        Call Initializing(m, vm, pim, ip, cmcp2, ccep2)
    End Sub
End Class
