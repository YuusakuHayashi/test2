Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Forms

Public Class ProjectViewModel : Inherits BaseViewModel2
    Private _IconFileName As String
    Public Property IconFileName As String
        Get
            Return _IconFileName
        End Get
        Set(value As String)
            _IconFileName = value
        End Set
    End Property

    Private _Icon As BitmapImage
    <JsonIgnore>
    Public Property [Icon] As BitmapImage
        Get
            Return _Icon
        End Get
        Set(value As BitmapImage)
            _Icon = value
        End Set
    End Property

    Private _Name As String
    Public Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            _Name = value
        End Set
    End Property

    Private _Kind As String
    Public Property Kind As String
        Get
            Return _Kind
        End Get
        Set(value As String)
            _Kind = value
        End Set
    End Property

    Private _LastUpdate As String
    Public Property LastUpdate As String
        Get
            Return _LastUpdate
        End Get
        Set(value As String)
            _LastUpdate = value
        End Set
    End Property

    Private _RootDirectoryName As String
    Public Property RootDirectoryName As String
        Get
            Return _RootDirectoryName
        End Get
        Set(value As String)
            _RootDirectoryName = value
        End Set
    End Property

    Private _DirectoryName As String
    Public Property DirectoryName As String
        Get
            Return _DirectoryName
        End Get
        Set(value As String)
            _DirectoryName = value
        End Set
    End Property

    ' プロジェクトのモデルファイルです
    Private _ModelFileName As String
    Public Property ModelFileName As String
        Get
            Return Me._ModelFileName
        End Get
        Set(value As String)
            Me._ModelFileName = value
            RaisePropertyChanged("ModelFileName")
        End Set
    End Property

    Private _ProjectInfoFileName As String
    Public Property ProjectInfoFileName As String
        Get
            Return Me._ProjectInfoFileName
        End Get
        Set(value As String)
            Me._ProjectInfoFileName = value
            RaisePropertyChanged("ProjectInfoFileName")
        End Set
    End Property

    Private _ViewModelFileName As String
    Public Property ViewModelFileName As String
        Get
            Return Me._ViewModelFileName
        End Get
        Set(value As String)
            Me._ViewModelFileName = value
            RaisePropertyChanged("ViewModelFileName")
        End Set
    End Property

    ' コマンドプロパティ（Ｐｒｏｊｅｃｔをピン止め）
    '---------------------------------------------------------------------------------------------'
    'Private _FixProjectCommand As ICommand
    '<JsonIgnore>
    'Public ReadOnly Property FixProjectCommand As ICommand
    '    Get
    '        If Me._FixProjectCommand Is Nothing Then
    '            Me._FixProjectCommand = New DelegateCommand With {
    '                .ExecuteHandler = AddressOf _FixProjectCommandExecute,
    '                .CanExecuteHandler = AddressOf _FixProjectCommandCanExecute
    '            }
    '            Return Me._FixProjectCommand
    '        Else
    '            Return Me._FixProjectCommand
    '        End If
    '    End Get
    'End Property

    'Private Sub _CheckFixProjectCommandEnabled()
    '    Dim b As Boolean : b = True
    '    Me._FixProjectCommandEnableFlag = b
    'End Sub

    'Private __FixProjectCommandEnableFlag As Boolean
    'Private Property _FixProjectCommandEnableFlag As Boolean
    '    Get
    '        Return Me.__FixProjectCommandEnableFlag
    '    End Get
    '    Set(value As Boolean)
    '        Me.__FixProjectCommandEnableFlag = value
    '        RaisePropertyChanged("_FixProjectCommandEnableFlag")
    '        CType(FixProjectCommand, DelegateCommand).RaiseCanExecuteChanged()
    '    End Set
    'End Property

    'Private Sub _FixProjectCommandExecute(ByVal parameter As Object)
    '    Call DelegateEventListener.Instance.RaiseFixProjectRequested(Me)
    'End Sub

    'Private Function _FixProjectCommandCanExecute(ByVal parameter As Object) As Boolean
    '    Return Me._FixProjectCommandEnableFlag
    'End Function
    '---------------------------------------------------------------------------------------------'

    ' コマンドプロパティ（Ｐｒｏｊｅｃｔのピン解除）
    '---------------------------------------------------------------------------------------------'
    'Private _RemoveFixedProjectCommand As ICommand
    '<JsonIgnore>
    'Public ReadOnly Property RemoveFixedProjectCommand As ICommand
    '    Get
    '        If Me._RemoveFixedProjectCommand Is Nothing Then
    '            Me._RemoveFixedProjectCommand = New DelegateCommand With {
    '                .ExecuteHandler = AddressOf _RemoveFixedProjectCommandExecute,
    '                .CanExecuteHandler = AddressOf _RemoveFixedProjectCommandCanExecute
    '            }
    '            Return Me._RemoveFixedProjectCommand
    '        Else
    '            Return Me._RemoveFixedProjectCommand
    '        End If
    '    End Get
    'End Property

    'Private Sub _CheckRemoveFixedProjectCommandEnabled()
    '    Dim b As Boolean : b = True
    '    Me._RemoveFixedProjectCommandEnableFlag = b
    'End Sub

    'Private __RemoveFixedProjectCommandEnableFlag As Boolean
    'Private Property _RemoveFixedProjectCommandEnableFlag As Boolean
    '    Get
    '        Return Me.__RemoveFixedProjectCommandEnableFlag
    '    End Get
    '    Set(value As Boolean)
    '        Me.__RemoveFixedProjectCommandEnableFlag = value
    '        RaisePropertyChanged("_RemoveFixedProjectCommandEnableFlag")
    '        CType(RemoveFixedProjectCommand, DelegateCommand).RaiseCanExecuteChanged()
    '    End Set
    'End Property

    'Private Sub _RemoveFixedProjectCommandExecute(ByVal parameter As Object)
    '    Call DelegateEventListener.Instance.RaiseRemoveFixedProjectRequested(Me)
    'End Sub

    'Private Function _RemoveFixedProjectCommandCanExecute(ByVal parameter As Object) As Boolean
    '    Return Me._RemoveFixedProjectCommandEnableFlag
    'End Function
    '---------------------------------------------------------------------------------------------'

    ' コマンドプロパティ（ファイル編集）
    '---------------------------------------------------------------------------------------------'
    Private _EditFileCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property EditFileCommand As ICommand
        Get
            If Me._EditFileCommand Is Nothing Then
                Me._EditFileCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _EditFileCommandExecute,
                    .CanExecuteHandler = AddressOf _EditFileCommandCanExecute
                }
                Return Me._EditFileCommand
            Else
                Return Me._EditFileCommand
            End If
        End Get
    End Property

    Private Sub _CheckEditFileCommandEnabled()
        Dim b As Boolean : b = True
        Me._EditFileCommandEnableFlag = b
    End Sub

    Private __EditFileCommandEnableFlag As Boolean
    Private Property _EditFileCommandEnableFlag As Boolean
        Get
            Return Me.__EditFileCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__EditFileCommandEnableFlag = value
            RaisePropertyChanged("_EditFileCommandEnableFlag")
            CType(EditFileCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _EditFileCommandExecute(ByVal parameter As Object)
        Dim f = CType(parameter, String)
        Dim psi As New System.Diagnostics.ProcessStartInfo With {
            .FileName = "notepad.exe",
            .Arguments = $"{f}"
        }
        Dim p = System.Diagnostics.Process.Start(psi)
    End Sub

    Private Function _EditFileCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._EditFileCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    ' コマンドプロパティ（ファイル選択）
    '---------------------------------------------------------------------------------------------'
    Private _InputFileDialogCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property InputFileDialogCommand As ICommand
        Get
            If Me._InputFileDialogCommand Is Nothing Then
                Me._InputFileDialogCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _InputFileDialogCommandExecute,
                    .CanExecuteHandler = AddressOf _InputFileDialogCommandCanExecute
                }
                Return Me._InputFileDialogCommand
            Else
                Return Me._InputFileDialogCommand
            End If
        End Get
    End Property

    Private Sub _CheckInputFileDialogCommandEnabled()
        Dim b As Boolean : b = True
        Me._InputFileDialogCommandEnableFlag = b
    End Sub

    Private __InputFileDialogCommandEnableFlag As Boolean
    Private Property _InputFileDialogCommandEnableFlag As Boolean
        Get
            Return Me.__InputFileDialogCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__InputFileDialogCommandEnableFlag = value
            RaisePropertyChanged("_InputFileDialogCommandEnableFlag")
            CType(InputFileDialogCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _InputFileDialogCommandExecute(ByVal parameter As Object)
        Dim f As String
        Dim t As String
        Dim fil = "JSONファイル(*.json)|*.json"
        Select Case parameter
            Case Me.ModelFileName
                f = "Model.json"
                t = "Select Model File"
            Case Me.ViewModelFileName
                f = "ViewModel.json"
                t = "Select ViewModel File"
            Case Me.ProjectInfoFileName
                f = "ProjectInfo.json"
                t = "Select Project Information File"
            Case Me.IconFileName
                f = "Project.ico"
                t = "プロジェクトアイコンの選択"
                fil = "すべてのファイル(*.*)|*.*"
        End Select
        Dim ofd As New OpenFileDialog With {
            .FileName = f,
            .InitialDirectory = Me.DirectoryName,
            .Filter = fil,
            .FilterIndex = 1,
            .Title = t
        }
        If ofd.ShowDialog() = DialogResult.OK Then
            Me.ModelFileName = ofd.FileName
        End If
    End Sub

    Private Function _InputFileDialogCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._InputFileDialogCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

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
        Dim f = CType(parameter, String)
        Dim fbd As New FolderBrowserDialog With {
            .Description = "プロジェクトフォルダの指定",
            .RootFolder = Me.RootDirectoryName,
            .SelectedPath = Me.DirectoryName,
            .ShowNewFolderButton = True
        }
        If fbd.ShowDialog() = DialogResult.OK Then
            Me.DirectoryName = fbd.SelectedPath
        End If
    End Sub

    Private Function _FolderBrowserDialogCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._FolderBrowserDialogCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Sub _ViewInitializing()
        With AppInfo.ProjectInfo
            Me.IconFileName = .IconFileName
            Me.[Icon] = .[Icon]
            Me.Name = .Name
            Me.Kind = .Kind
            Me.LastUpdate = .LastUpdate
            Me.RootDirectoryName = .RootDirectoryName
            Me.DirectoryName = .DirectoryName
            Me.ModelFileName = .ModelFileName
            Me.ViewModelFileName = .ViewModelFileName
            Me.ProjectInfoFileName = .ProjectInfoFileName
        End With
    End Sub

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel,
                                    ByRef vm As ViewModel)
        InitializeHandler _
            = AddressOf _ViewInitializing
        CheckCommandEnabledHandler _
            = [Delegate].Combine(
                New Action(AddressOf _CheckEditFileCommandEnabled),
                New Action(AddressOf _CheckInputFileDialogCommandEnabled),
                New Action(AddressOf _CheckFolderBrowserDialogCommandEnabled)
            )
        Call BaseInitialize(app, vm)
    End Sub
End Class
