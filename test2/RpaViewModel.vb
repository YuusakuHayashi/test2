Imports test2
Imports System.IO
Imports System.Windows.Forms

Public Class RpaViewModel : Inherits BaseViewModel2
    Private Const _INITIAL_ROOT_PROJECT_NAME = "ルートプロジェクト名を入力してください"
    Private Const _INITIAL_USER_PROJECT_NAME = "プロジェクト名を入力してください"

    Private Const DQ As String = """"

    Private ReadOnly Property _RootDirectoryName As String
        Get
            Return AppInfo.ProjectInfo.Model.Data.RootDirectoryName
        End Get
    End Property

    Private ReadOnly Property _UserDirectoryName As String
        Get
            Return AppInfo.ProjectInfo.Model.Data.UserDirectoryName
        End Get
    End Property

    Private _Rpa As RpaModel

    Private _RootProjectName As String
    Public Property RootProjectName As String
        Get
            Return Me._RootProjectName
        End Get
        Set(value As String)
            If Not Directory.Exists(Me._RootDirectoryName & "\" & value) Then
                Me.ErrorMessage = $"Error: ルートプロジェクト{DQ}{value}{DQ}は存在しません"
            Else
                Me.ErrorMessage = vbNullString
            End If
            Me._RootProjectName = value
            Me._Rpa.RootProjectName = value
            Call _RpaUpdate()
            RaisePropertyChanged("RootProjectName")
        End Set
    End Property

    Private _UserProjectName As String
    Public Property UserProjectName As String
        Get
            Return Me._UserProjectName
        End Get
        Set(value As String)
            If Not Directory.Exists(Me._UserDirectoryName & "\" & value) Then
                Me.ErrorMessage = $"Error: プロジェクト{DQ}{value}{DQ}は存在しません"
            Else
                Me.ErrorMessage = vbNullString
            End If
            Me._UserProjectName = value
            Me._Rpa.UserProjectName = value
            Call _RpaUpdate()
            RaisePropertyChanged("UserProjectName")
        End Set
    End Property

    Private _Status As Integer
    Public Property Status As Integer
        Get
            Return Me._Status
        End Get
        Set(value As Integer)
            Me._Status = value
            Me._Rpa.Status = value
            Call _RpaUpdate()
            Call _CheckConnectProjectCommandEnabled()
            RaisePropertyChanged("StatusText")
        End Set
    End Property

    Private _StatusText As String
    Public ReadOnly Property StatusText As String
        Get
            Select Case Me.Status
                Case 0
                    Me._StatusText = "未接続"
                Case 1
                    Me._StatusText = "接続済み"
            End Select
            Return Me._StatusText
        End Get
    End Property

    Private _ErrorMessage As String
    Public Property ErrorMessage As String
        Get
            Return Me._ErrorMessage
        End Get
        Set(value As String)
            Me._ErrorMessage = value
            Call _CheckConnectProjectCommandEnabled()
            RaisePropertyChanged("ErrorMessage")
        End Set
    End Property

    Private _RunParameter As String
    Public Property RunParameter As String
        Get
            Return Me._RunParameter
        End Get
        Set(value As String)
            Me._RunParameter = value
            Me._Rpa.RunParameter = value
            Call _RpaUpdate()
            RaisePropertyChanged("RunParameter")
        End Set
    End Property

    ' コマンドプロパティ（フォルダ選択）
    '---------------------------------------------------------------------------------------------'
    Private _FolderBrowserDialogCommand As ICommand
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
        Dim d = CType(parameter, String)
        Dim b = True
        Dim [dir] As String

        Select Case d
            Case Me.RootProjectName
                dsc = "ルートプロジェクトの指定"
                [dir] = Me._RootDirectoryName
                b = False
            Case Me.UserProjectName
                dsc = "プロジェクトの指定"
                [dir] = Me._UserDirectoryName
        End Select

        fbd = New FolderBrowserDialog With {
            .Description = dsc,
            .SelectedPath = $"{[dir]}",
            .ShowNewFolderButton = b
        }

        If fbd.ShowDialog() = DialogResult.OK Then
            Select Case d
                Case Me.RootProjectName
                    Me.RootProjectName = fbd.SelectedPath.Replace(Me._RootDirectoryName & "\", "")
                Case Me.UserProjectName
                    Me.UserProjectName = fbd.SelectedPath.Replace(Me._UserDirectoryName & "\", "")
            End Select
        End If
    End Sub

    Private Function _FolderBrowserDialogCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._FolderBrowserDialogCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    ' コマンドプロパティ（フォルダ選択）
    '---------------------------------------------------------------------------------------------'
    Private _ConnectProjectCommand As ICommand
    Public ReadOnly Property ConnectProjectCommand As ICommand
        Get
            If Me._ConnectProjectCommand Is Nothing Then
                Me._ConnectProjectCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ConnectProjectCommandExecute,
                    .CanExecuteHandler = AddressOf _ConnectProjectCommandCanExecute
                }
                Return Me._ConnectProjectCommand
            Else
                Return Me._ConnectProjectCommand
            End If
        End Get
    End Property

    Private Sub _CheckConnectProjectCommandEnabled()
        Dim b As Boolean : b = True
        If Not Me.ErrorMessage = vbNullString Then
            b = False
        End If
        If Me.Status > 0 Then
            b = False
        End If
        Me._ConnectProjectCommandEnableFlag = b
    End Sub

    Private __ConnectProjectCommandEnableFlag As Boolean
    Private Property _ConnectProjectCommandEnableFlag As Boolean
        Get
            Return Me.__ConnectProjectCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ConnectProjectCommandEnableFlag = value
            RaisePropertyChanged("_ConnectProjectCommandEnableFlag")
            CType(ConnectProjectCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _ConnectProjectCommandExecute(ByVal parameter As Object)
        Me.Status = 1
    End Sub

    Private Function _ConnectProjectCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ConnectProjectCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Sub _RpaUpdate()
        AppInfo.ProjectInfo.Model.Data.Rpas(
            AppInfo.ProjectInfo.Model.Data.Rpas.IndexOf(Me._Rpa)
        ) = Me._Rpa
    End Sub

    Public Sub _ViewInitializing()
        Dim b = False
        Dim [rpa] As RpaModel
        For Each r In AppInfo.ProjectInfo.Model.Data.Rpas
            If Not r.IsViewAssigned Then
                Me._Rpa = r
                Me.RootProjectName = r.RootProjectName
                Me.UserProjectName = r.UserProjectName
                Me.Status = r.Status
                Me.RunParameter = r.RunParameter
                r.IsViewAssigned = True
                b = True
                Exit For
            End If
        Next
        If Not b Then
            [rpa] = New RpaModel With {
                .Index = (AppInfo.ProjectInfo.Model.Data.GetRpaIndex()),
                .IsViewAssigned = True
            }
            AppInfo.ProjectInfo.Model.Data.Rpas.Add([rpa])
            Me._Rpa = AppInfo.ProjectInfo.Model.Data.Rpas(
                AppInfo.ProjectInfo.Model.Data.Rpas.IndexOf([rpa])
            )
            Me.RootProjectName = _INITIAL_ROOT_PROJECT_NAME
            Me.UserProjectName = _INITIAL_USER_PROJECT_NAME
            Me.RunParameter = "Run"
        End If
    End Sub

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        InitializeHandler = AddressOf _ViewInitializing
        CheckCommandEnabledHandler _
            = [Delegate].Combine(
                New Action(AddressOf _CheckFolderBrowserDialogCommandEnabled),
                New Action(AddressOf _CheckConnectProjectCommandEnabled)
            )
        Call BaseInitialize(app, vm)
    End Sub
End Class
