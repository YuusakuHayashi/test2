Imports test2
Imports System.Collections.ObjectModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Forms
Imports System.IO

Public Class RpaProjectViewModel : Inherits BaseViewModel2

    Private _RootDirectoryName As String
    Public Property RootDirectoryName As String
        Get
            If Me._RootDirectoryName Is Nothing Then
                Me._RootDirectoryName = AppInfo.ProjectInfo.DirectoryName
            End If
            Return Me._RootDirectoryName
        End Get
        Set(value As String)
            Me._RootDirectoryName = value
        End Set
    End Property

    Private _UserDirectoryName As String
    Public Property UserDirectoryName As String
        Get
            If Me._UserDirectoryName Is Nothing Then
                Me._UserDirectoryName = AppInfo.ProjectInfo.DirectoryName
            End If
            Return Me._UserDirectoryName
        End Get
        Set(value As String)
            Me._UserDirectoryName = value
        End Set
    End Property

    Private _Rpas As ObservableCollection(Of RpaModel)
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
        Dim f = CType(parameter, String)
        Dim pd = Directory.GetParent(f).FullName
        Dim fbd As New FolderBrowserDialog With {
            .Description = "ルートディレクトリの指定",
            .SelectedPath = $"{f}",
            .ShowNewFolderButton = True
        }
        If fbd.ShowDialog() = DialogResult.OK Then
            Me.RootDirectoryName = fbd.SelectedPath
        End If
    End Sub

    Private Function _FolderBrowserDialogCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._FolderBrowserDialogCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'


    Public Sub _ViewInitializing()
        With AppInfo.ProjectInfo.Model.Data
            Me.RootDirectoryName = .RootDirectoryName
            Me.UserDirectoryName = .UserDirectoryName
            Me.Rpas = .Rpas
        End With
    End Sub


    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        InitializeHandler = AddressOf _ViewInitializing
        CheckCommandEnabledHandler _
            = [Delegate].Combine(
                New Action(AddressOf _CheckFolderBrowserDialogCommandEnabled)
            )
        Call BaseInitialize(app, vm)
    End Sub
End Class
