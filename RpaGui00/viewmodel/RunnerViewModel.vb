Imports System.Collections.ObjectModel
Imports Newtonsoft.Json

Public Class RunnerViewModel : Inherits ExecutableMenuBase(Of RunnerViewModel)
    <JsonIgnore>
    Public ReadOnly Property RunnerFileName As String
        Get
            Dim fil As String = vbNullString
            If Data IsNot Nothing Then
                If Data.Project IsNot Nothing Then
                    fil = $"{Data.Project.MyDirectory}\runner.json"
                End If
            End If
            Return fil
        End Get
    End Property

    Private _ParametersText As String
    <JsonIgnore>
    Public Property ParametersText As String
        Get
            Return Me._ParametersText
        End Get
        Set(value As String)
            Me._ParametersText = value
            RaisePropertyChanged("ParametersText")
        End Set
    End Property

    Private _StatusCode As Integer
    <JsonIgnore>
    Public Property StatusCode As Integer
        Get
            Return Me._StatusCode
        End Get
        Set(value As Integer)
            Me._StatusCode = value
            RaisePropertyChanged("StatusCode")
        End Set
    End Property

    Private _ParametersTexts As ObservableCollection(Of String)
    Public Property ParametersTexts As ObservableCollection(Of String)
        Get
            If Me._ParametersTexts Is Nothing Then
                Me._ParametersTexts = New ObservableCollection(Of String)
            End If
            Return Me._ParametersTexts
        End Get
        Set(value As ObservableCollection(Of String))
            Me._ParametersTexts = value
            RaisePropertyChanged("ParametersTexts")
        End Set
    End Property

    Private _IndexOfParametersText As Integer
    <JsonIgnore>
    Public Property IndexOfParametersText As Integer
        Get
            Return Me._IndexOfParametersText
        End Get
        Set(value As Integer)
            Me._IndexOfParametersText = value
            ' ここに書くのは気持ち悪いけど、面倒なので一旦ここに記述
            Me.ParametersText = Me.ParametersTexts(value)
            RaisePropertyChanged("IndexOfParametersText")
        End Set
    End Property

    Private _RunRobotIcon As BitmapImage
    <JsonIgnore>
    Public Property RunRobotIcon As BitmapImage
        Get
            If Me._RunRobotIcon Is Nothing Then
                Me._RunRobotIcon = CreateIcon($"{Controller.SystemDllDirectory}\runrobot.png")
            End If
            Return _RunRobotIcon
        End Get
        Set(value As BitmapImage)
            Me._RunRobotIcon = value
            RaisePropertyChanged("RunRobotIcon")
        End Set
    End Property

    Private _CheckRobotIcon As BitmapImage
    <JsonIgnore>
    Public Property CheckRobotIcon As BitmapImage
        Get
            If Me._CheckRobotIcon Is Nothing Then
                Me._CheckRobotIcon = CreateIcon($"{Controller.SystemDllDirectory}\checkrobot.png")
            End If
            Return _CheckRobotIcon
        End Get
        Set(value As BitmapImage)
            Me._CheckRobotIcon = value
            RaisePropertyChanged("CheckRobotIcon")
        End Set
    End Property

    Private _ReloadMyRobotReadMeIcon As BitmapImage
    <JsonIgnore>
    Public Property ReloadMyRobotReadMeIcon As BitmapImage
        Get
            If Me._ReloadMyRobotReadMeIcon Is Nothing Then
                Me._ReloadMyRobotReadMeIcon = CreateIcon($"{Controller.SystemDllDirectory}\reloadmyrobotreadme.png")
            End If
            Return _ReloadMyRobotReadMeIcon
        End Get
        Set(value As BitmapImage)
            Me._ReloadMyRobotReadMeIcon = value
            RaisePropertyChanged("ReloadMyRobotReadMeIcon")
        End Set
    End Property

    Private _ProcessTime As Integer
    <JsonIgnore>
    Public Property ProcessTime As Integer
        Get
            Return Me._ProcessTime
        End Get
        Set(value As Integer)
            Me._ProcessTime = value
            RaisePropertyChanged("ProcessTime")
        End Set
    End Property

    Private _PreProcessTime As Integer
    <JsonIgnore>
    Public Property PreProcessTime As Integer
        Get
            Return Me._PreProcessTime
        End Get
        Set(value As Integer)
            Me._PreProcessTime = value
            RaisePropertyChanged("PreProcessTime")
        End Set
    End Property

    Private _RunProcessTime As Integer
    <JsonIgnore>
    Public Property RunProcessTime As Integer
        Get
            Return Me._RunProcessTime
        End Get
        Set(value As Integer)
            Me._RunProcessTime = value
            RaisePropertyChanged("RunProcessTime")
        End Set
    End Property

    Private _PreRunProcessTime As Integer
    Public Property PreRunProcessTime As Integer
        Get
            Return Me._PreRunProcessTime
        End Get
        Set(value As Integer)
            Me._PreRunProcessTime = value
            RaisePropertyChanged("PreRunProcessTime")
        End Set
    End Property

    Private _CheckProcessTime As Integer
    <JsonIgnore>
    Public Property CheckProcessTime As Integer
        Get
            Return Me._CheckProcessTime
        End Get
        Set(value As Integer)
            Me._CheckProcessTime = value
            RaisePropertyChanged("CheckProcessTime")
        End Set
    End Property

    Private _PreCheckProcessTime As Integer
    Public Property PreCheckProcessTime As Integer
        Get
            Return Me._PreCheckProcessTime
        End Get
        Set(value As Integer)
            Me._PreCheckProcessTime = value
            RaisePropertyChanged("PreCheckProcessTime")
        End Set
    End Property

    Private _IsProcessDone As Boolean
    <JsonIgnore>
    Public Property IsProcessDone As Boolean
        Get
            Return Me._IsProcessDone
        End Get
        Set(value As Boolean)
            Me._IsProcessDone = value
            RaisePropertyChanged("IsProcessDone")
        End Set
    End Property

    Private _ProjectName As String
    <JsonIgnore>
    Public Property ProjectName As String
        Get
            Return Me._ProjectName
        End Get
        Set(value As String)
            Me._ProjectName = value
            RaisePropertyChanged("ProjectName")
        End Set
    End Property

    Private _RobotName As String
    <JsonIgnore>
    Public Property RobotName As String
        Get
            Return Me._RobotName
        End Get
        Set(value As String)
            Me._RobotName = value
            RaisePropertyChanged("RobotName")
        End Set
    End Property

    Private _RobotAlias As String
    <JsonIgnore>
    Public Property RobotAlias As String
        Get
            Return Me._RobotAlias
        End Get
        Set(value As String)
            Me._RobotAlias = value
            RaisePropertyChanged("RobotAlias")
        End Set
    End Property

    Private _RootRobotDirectory As String
    <JsonIgnore>
    Public Property RootRobotDirectory As String
        Get
            Return Me._RootRobotDirectory
        End Get
        Set(value As String)
            Me._RootRobotDirectory = value
            RaisePropertyChanged("RootRobotDirectory")
        End Set
    End Property

    Private _MyRobotDirectory As String
    <JsonIgnore>
    Public Property MyRobotDirectory As String
        Get
            Return Me._MyRobotDirectory
        End Get
        Set(value As String)
            Me._MyRobotDirectory = value
            RaisePropertyChanged("MyRobotDirectory")
        End Set
    End Property

    Private _MyRobotReadMe As String
    <JsonIgnore>
    Public Property MyRobotReadMe As String
        Get
            If String.IsNullOrEmpty(Me._MyRobotReadMe) Then
                If Data IsNot Nothing Then
                    If Data.Project IsNot Nothing Then
                        Me._MyRobotReadMe = Data.Project.MyRobotReadMe
                    End If
                End If
            End If
            Return Me._MyRobotReadMe
        End Get
        Set(value As String)
            Me._MyRobotReadMe = value
            RaisePropertyChanged("MyRobotReadMe")
        End Set
    End Property

    '---------------------------------------------------------------------------------------------'
    Private _CheckRobotCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property CheckRobotCommand As ICommand
        Get
            If Me._CheckRobotCommand Is Nothing Then
                Me._CheckRobotCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf CheckRobotCommandExecute,
                    .CanExecuteHandler = AddressOf CheckRobotCommandCanExecute
                }
            End If
            Return Me._CheckRobotCommand
        End Get
    End Property

    Private _IsCheckRobotCommandEnabled As Boolean
    <JsonIgnore>
    Public Property IsCheckRobotCommandEnabled As Boolean
        Get
            Return Me._IsCheckRobotCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsCheckRobotCommandEnabled = value
            RaisePropertyChanged("IsCheckRobotCommandEnabled")
            CType(CheckRobotCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _CheckRobotCommandExecute(ByVal parameter As Object)
        Me.StatusCode = 1
        Data.System.ExecuteMode = Rpa00.RpaSystem.ExecuteModeNumber.RpaGui
        Data.System.GuiCommandText = $"checkrobot"
        Call Data.System.CuiLoop(Data)
    End Sub

    Private Async Sub CheckRobotCommandExecute(ByVal parameter As Object)
        Me.IsCheckRobotCommandEnabled = False
        Dim t As Task = Task.Run(
            Sub()
                Call _CheckRobotCommandExecute(parameter)
            End Sub
        )
        Dim t2 As Task = Task.Run(
            Sub()
                Call CountProcessTime()
            End Sub
        )
        Await Task.WhenAll(t, t2)
        Me.IsCheckRobotCommandEnabled = True
        Me.StatusCode = 2
        Me.PreCheckProcessTime = Me.ProcessTime
        Me.ProcessTime = 0
    End Sub

    Private Function CheckRobotCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsCheckRobotCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    Private _RunRobotCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property RunRobotCommand As ICommand
        Get
            If Me._RunRobotCommand Is Nothing Then
                Me._RunRobotCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf RunRobotCommandExecute,
                    .CanExecuteHandler = AddressOf RunRobotCommandCanExecute
                }
            End If
            Return Me._RunRobotCommand
        End Get
    End Property

    Private _IsRunRobotCommandEnabled As Boolean
    <JsonIgnore>
    Public Property IsRunRobotCommandEnabled As Boolean
        Get
            Return Me._IsRunRobotCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsRunRobotCommandEnabled = value
            RaisePropertyChanged("IsRunRobotCommandEnabled")
            CType(RunRobotCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _RunRobotCommandExecute(ByVal parameter As Object)
        Me.StatusCode = 1
        Data.System.ExecuteMode = Rpa00.RpaSystem.ExecuteModeNumber.RpaGui
        Data.System.GuiCommandText = $"runrobot {Me.ParametersText}"
        Call Data.System.CuiLoop(Data)
    End Sub

    Private Async Sub RunRobotCommandExecute(ByVal parameter As Object)
        Me.IsRunRobotCommandEnabled = False
        Dim t As Task = Task.Run(
            Sub()
                Call _RunRobotCommandExecute(parameter)
            End Sub
        )
        Dim t2 As Task = Task.Run(
            Sub()
                Me.PreProcessTime = Me.PreRunProcessTime
                Call CountProcessTime()
            End Sub
        )
        Await Task.WhenAll(t, t2)
        Me.IsRunRobotCommandEnabled = True
        Me.StatusCode = 2
        Me.PreRunProcessTime = Me.ProcessTime
        Me.ProcessTime = 0
        Dim i As Integer = AddParametersText()
        Call Save(Me.RunnerFileName, Me)
    End Sub

    Private Sub CountProcessTime()
        Do
            Threading.Thread.Sleep(1000)
            Me.ProcessTime += 1
            If Me.ProcessTime > Me.PreProcessTime Then
                Me.ProcessTime = 1
            End If
        Loop Until Not Data.System.IsRunTime
    End Sub

    Private Function RunRobotCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsRunRobotCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    Private _ReloadMyRobotReadMeCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property ReloadMyRobotReadMeCommand As ICommand
        Get
            If Me._ReloadMyRobotReadMeCommand Is Nothing Then
                Me._ReloadMyRobotReadMeCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf ReloadMyRobotReadMeCommandExecute,
                    .CanExecuteHandler = AddressOf ReloadMyRobotReadMeCommandCanExecute
                }
            End If
            Return Me._ReloadMyRobotReadMeCommand
        End Get
    End Property

    Private _IsReloadMyRobotReadMeCommandEnabled As Boolean
    <JsonIgnore>
    Public Property IsReloadMyRobotReadMeCommandEnabled As Boolean
        Get
            Return Me._IsReloadMyRobotReadMeCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsReloadMyRobotReadMeCommandEnabled = value
            RaisePropertyChanged("IsReloadMyRobotReadMeCommandEnabled")
            CType(ReloadMyRobotReadMeCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub ReloadMyRobotReadMeCommandExecute(ByVal parameter As Object)
        Me.MyRobotReadMe = vbNullString
    End Sub

    Private Function ReloadMyRobotReadMeCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsReloadMyRobotReadMeCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Function AddParametersText() As Integer
        If Not Me.ParametersTexts.Contains(Me.ParametersText) Then
            Me.ParametersTexts = Rpa00.RpaModule.Push(Of String, ObservableCollection(Of String))(Me.ParametersText, Me.ParametersTexts)
        End If
        Return 0
    End Function

    Private Function CreateIcon(ByVal filename As String) As BitmapImage
        Dim bi = New BitmapImage
        bi.BeginInit()
        bi.UriSource = New Uri(filename, UriKind.Absolute)
        bi.EndInit()
        Return bi
    End Function

    Public Overrides Sub Initialize(ByRef dat As Object)
        MyBase.Initialize(dat)
        If Data IsNot Nothing Then
            If Data.Project IsNot Nothing Then
                Me.StatusCode = 0
                Me.ProjectName = Data.Project.ProjectName
                Me.RobotName = Data.Project.RobotName
                Me.RobotAlias = Data.Project.RobotAlias
                Me.MyRobotDirectory = Data.Project.MyRobotDirectory
                Me.RootRobotDirectory = Data.Project.RootRobotDirectory
                Me.IsRunRobotCommandEnabled = True
                Me.IsCheckRobotCommandEnabled = True
                Me.ProcessTime = 0
                Me.PreProcessTime = 60
                Me.PreRunProcessTime = 60
                Me.PreCheckProcessTime = 60
                Me.IsReloadMyRobotReadMeCommandEnabled = True
                'AddHandler Me.ParametersTexts.CollectionChanged, AddressOf UpdateParametersTexts
            End If
        End If
    End Sub

    Sub New()
    End Sub
End Class
