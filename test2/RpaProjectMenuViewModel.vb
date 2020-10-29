Imports test2
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class RpaProjectMenuViewModel : Inherits BaseViewModel2

    Private _StartupIcon As BitmapImage
    Public Property StartupIcon As BitmapImage
        Get
            Return Me._StartupIcon
        End Get
        Set(value As BitmapImage)
            Me._StartupIcon = value
            RaisePropertyChanged("StartupIcon")
        End Set
    End Property

    Private _DownloadIcon As BitmapImage
    Public Property DownloadIcon As BitmapImage
        Get
            Return Me._DownloadIcon
        End Get
        Set(value As BitmapImage)
            Me._DownloadIcon = value
            RaisePropertyChanged("DownloadIcon")
        End Set
    End Property

    ' コマンドプロパティ（フォルダ選択）
    '---------------------------------------------------------------------------------------------'
    Private _RpaLaunchCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property RpaLaunchCommand As ICommand
        Get
            If Me._RpaLaunchCommand Is Nothing Then
                Me._RpaLaunchCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _RpaLaunchCommandExecute,
                    .CanExecuteHandler = AddressOf _RpaLaunchCommandCanExecute
                }
                Return Me._RpaLaunchCommand
            Else
                Return Me._RpaLaunchCommand
            End If
        End Get
    End Property

    Private Sub _CheckRpaLaunchCommandEnabled()
        Dim b As Boolean : b = True
        Me._RpaLaunchCommandEnableFlag = b
    End Sub

    Private __RpaLaunchCommandEnableFlag As Boolean
    Private Property _RpaLaunchCommandEnableFlag As Boolean
        Get
            Return Me.__RpaLaunchCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__RpaLaunchCommandEnableFlag = value
            RaisePropertyChanged("_RpaLaunchCommandEnableFlag")
            CType(RpaLaunchCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _RpaLaunchCommandExecute(ByVal parameter As Object)
        Dim v As ViewItemModel
        Dim rpavm As New RpaViewModel
        rpavm.Initialize(AppInfo, ViewModel)
        v = New ViewItemModel With {
            .Content = rpavm,
            .Name = "BrankRpa" & _GetIndex().ToString(),
            .FrameType = MultiViewModel.MAIN_FRAME,
            .LayoutType = ViewModel.MULTI_VIEW,
            .ViewType = MultiViewModel.TAB_VIEW,
            .OpenState = True,
            .ModelName = rpavm.GetType.Name
        }
        Call AddViewItem(v)
        Call AddView(v)
    End Sub

    Private Function _RpaLaunchCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._RpaLaunchCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Function _GetIndex() As Integer
        Dim i = 0
        Dim b = False
        With AppInfo.ProjectInfo.Model.Data
            Do Until True = False
                b = False
                For Each rpa In .Rpas
                    If i = rpa.Index Then
                        i += 1
                        b = True
                        Exit For
                    End If
                Next
                If Not b Then
                    Exit Do
                End If
            Loop
        End With
        _GetIndex = i
    End Function

    Public Sub _ViewInitialize()
        Dim f As Func(Of String, BitmapImage)
        f = Function(ByVal img As String) As BitmapImage
                Dim bi = New BitmapImage
                bi.BeginInit()
                bi.UriSource = New Uri(
                    img,
                    UriKind.Absolute
                )
                bi.EndInit()
                Return bi
            End Function

        Me.StartupIcon = f(AppDirectoryModel.AppImageDirectory & "\RpaProject\startup.png")
        Me.DownloadIcon = f(AppDirectoryModel.AppImageDirectory & "\RpaProject\download.png")
    End Sub

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        InitializeHandler = AddressOf _ViewInitialize
        CheckCommandEnabledHandler = [Delegate].Combine(
            New Action(AddressOf _CheckRpaLaunchCommandEnabled)
        )
        BaseInitialize(app, vm)
    End Sub
End Class
