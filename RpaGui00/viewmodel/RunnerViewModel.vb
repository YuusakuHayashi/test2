﻿Imports System.ComponentModel
Imports System.IO
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Newtonsoft.Json

Public Class RunnerViewModel : Inherits ControllerViewModelBase(Of RunnerViewModel)
    Private Const RUNCOMMAND = "runrobot"
    Private Const CHKCOMMAND = "checkrobot"

    Private ReadOnly Property ProjectRunnerJsonFileName As String
        Get
            Dim fil As String = vbNullString
            If ViewController.Data IsNot Nothing Then
                If ViewController.Data.Project IsNot Nothing Then
                    fil = $"{ViewController.Data.Project.MyDirectory}\runner.json"
                End If
            End If
            Return fil
        End Get
    End Property

    Private ReadOnly Property RobotRunnerJsonFileName As String
        Get
            Dim fil As String = vbNullString
            If ViewController.Data IsNot Nothing Then
                If ViewController.Data.Project IsNot Nothing Then
                    fil = $"{ViewController.Data.Project.MyRobotDirectory}\runner.json"
                End If
            End If
            Return fil
        End Get
    End Property

    Private ReadOnly Property RunnerSaveIconFileName As String
        Get
            Return $"{Controller.SystemDllDirectory}\runnersave.png"
        End Get
    End Property
    Private _RunnerSaveIcon As BitmapImage
    Public ReadOnly Property RunnerSaveIcon As BitmapImage
        Get
            If Me._RunnerSaveIcon Is Nothing Then
                Me._RunnerSaveIcon = RpaGuiModule.CreateIcon(Me.RunnerSaveIconFileName)
            End If
            Return Me._RunnerSaveIcon
        End Get
    End Property

    Private ReadOnly Property RunnerRunRobotIconFileName As String
        Get
            Return $"{Controller.SystemDllDirectory}\runnerrunrobot.png"
        End Get
    End Property
    Private _RunnerRunRobotIcon As BitmapImage
    Public ReadOnly Property RunnerRunRobotIcon As BitmapImage
        Get
            If Me._RunnerRunRobotIcon Is Nothing Then
                Me._RunnerRunRobotIcon = RpaGuiModule.CreateIcon(Me.RunnerRunRobotIconFileName)
            End If
            Return Me._RunnerRunRobotIcon
        End Get
    End Property
    Private _RunRobotCommandToolTip As String
    Public Property RunRobotCommandToolTip As String
        Get
            If String.IsNullOrEmpty(Me._RunRobotCommandToolTip) Then
                Me._RunRobotCommandToolTip = Me.RunRobotCommandDatas(0).CommandText
            End If
            Return Me._RunRobotCommandToolTip
        End Get
        Set(value As String)
            Me._RunRobotCommandToolTip = value
            RaisePropertyChanged("RunRobotCommandToolTip")
        End Set
    End Property

    Private ReadOnly Property RunnerCheckRobotIconFileName As String
        Get
            Return $"{Controller.SystemDllDirectory}\runnercheckrobot.png"
        End Get
    End Property
    Private _RunnerCheckRobotIcon As BitmapImage
    Public ReadOnly Property RunnerCheckRobotIcon As BitmapImage
        Get
            If Me._RunnerCheckRobotIcon Is Nothing Then
                Me._RunnerCheckRobotIcon = RpaGuiModule.CreateIcon(Me.RunnerCheckRobotIconFileName)
            End If
            Return Me._RunnerCheckRobotIcon
        End Get
    End Property
    Private _CheckRobotCommandToolTip As String
    Public Property CheckRobotCommandToolTip As String
        Get
            If String.IsNullOrEmpty(Me._CheckRobotCommandToolTip) Then
                Me._CheckRobotCommandToolTip = Me.CheckRobotCommandDatas(0).CommandText
            End If
            Return Me._CheckRobotCommandToolTip
        End Get
        Set(value As String)
            Me._CheckRobotCommandToolTip = value
            RaisePropertyChanged("CheckRobotCommandToolTip")
        End Set
    End Property

    Private ReadOnly Property ReloadMyRobotReadMeIconFileName As String
        Get
            Return $"{Controller.SystemDllDirectory}\reloadmyrobotreadme.png"
        End Get
    End Property
    Private _ReloadMyRobotReadMeIcon As BitmapImage
    Public ReadOnly Property ReloadMyRobotReadMeIcon As BitmapImage
        Get
            If Me._ReloadMyRobotReadMeIcon Is Nothing Then
                Me._ReloadMyRobotReadMeIcon = RpaGuiModule.CreateIcon(Me.ReloadMyRobotReadMeIconFileName)
            End If
            Return Me._ReloadMyRobotReadMeIcon
        End Get
    End Property

    Private ReadOnly Property SaveMyRobotReadMeIconFileName As String
        Get
            Return $"{Controller.SystemDllDirectory}\savemyrobotreadme.png"
        End Get
    End Property
    Private _SaveMyRobotReadMeIcon As BitmapImage
    Public ReadOnly Property SaveMyRobotReadMeIcon As BitmapImage
        Get
            If Me._SaveMyRobotReadMeIcon Is Nothing Then
                Me._SaveMyRobotReadMeIcon = RpaGuiModule.CreateIcon(Me.SaveMyRobotReadMeIconFileName)
            End If
            Return Me._SaveMyRobotReadMeIcon
        End Get
    End Property

    Private _GuiCommand As String
    Private Property GuiCommand As String
        Get
            Return Me._GuiCommand
        End Get
        Set(value As String)
            Me._GuiCommand = value
            RaisePropertyChanged("GuiCommand")
        End Set
    End Property

    ' メニューの引数ありで起動の際に指定するコマンドテキスト
    Private _RunRobotCommandText As String
    <JsonIgnore>
    Public Property RunRobotCommandText As String
        Get
            Return Me._RunRobotCommandText
        End Get
        Set(value As String)
            Me._RunRobotCommandText = value
            RaisePropertyChanged("RunRobotCommandText")
        End Set
    End Property

    ' メニューの引数ありで起動の際に指定するコマンドテキスト
    Private _CheckRobotCommandText As String
    <JsonIgnore>
    Public Property CheckRobotCommandText As String
        Get
            Return Me._CheckRobotCommandText
        End Get
        Set(value As String)
            Me._CheckRobotCommandText = value
            RaisePropertyChanged("CheckRobotCommandText")
        End Set
    End Property

    Private _MyRobotReadMeContentVisibility As Visibility
    <JsonIgnore>
    Public Property MyRobotReadMeContentVisibility As Visibility
        Get
            If Me._MyRobotReadMeContentVisibility = Nothing Then
                Me._MyRobotReadMeContentVisibility = Visibility.Visible
            End If
            Return Me._MyRobotReadMeContentVisibility
        End Get
        Set(value As Visibility)
            Me._MyRobotReadMeContentVisibility = value
            RaisePropertyChanged("MyRobotReadMeContentVisibility")
        End Set
    End Property

    Private IsRunRobotCommandDatasDirectCast As Boolean
    ' 非常に複雑だが、RpaModule.Push などによって直接セットされる場合があり
    ' 直接セットされると、RunRobotCommandIndexも新たにセットされ、
    ' イベントハンドラが動いてしまう（ConsoleOutput.GuiCommandTextを更新させるための挙動）
    ' 直接セットする場合は、イベントハンドラは動かしたくないので、
    ' IsRunRobotCommandDataDirectCastによって制御している
    Private _RunRobotCommandDatas As ObservableCollection(Of CommandData)
    Public Property RunRobotCommandDatas As ObservableCollection(Of CommandData)
        Get
            Return Me._RunRobotCommandDatas
        End Get
        Set(value As ObservableCollection(Of CommandData))
            Me.IsRunRobotCommandDatasDirectCast = True
            Me._RunRobotCommandDatas = value
            RaisePropertyChanged("RunRobotCommandDatas")
            Me.IsRunRobotCommandDatasDirectCast = False
        End Set
    End Property

    Private _RunRobotCommandIndex As Integer
    <JsonIgnore>
    Public Property RunRobotCommandIndex As Integer
        Get
            Return Me._RunRobotCommandIndex
        End Get
        Set(value As Integer)
            If Me.IsRunRobotCommandDatasDirectCast Then
                Me._RunRobotCommandIndex = -1
            Else
                Me._RunRobotCommandIndex = value
            End If
            RaisePropertyChanged("RunRobotCommandIndex")
        End Set
    End Property

    Private IsCheckRobotCommandDatasDirectCast As Boolean
    ' 非常に複雑だが、RpaModule.Push などによって直接セットされる場合があり
    ' 直接セットされると、CheckRobotCommandIndexも新たにセットされ、
    ' イベントハンドラが動いてしまう（ConsoleOutput.GuiCommandTextを更新させるための挙動）
    ' 直接セットする場合は、イベントハンドラは動かしたくないので、
    ' IsCheckRobotCommandDataDirectCastによって制御している
    Private _CheckRobotCommandDatas As ObservableCollection(Of CommandData)
    Public Property CheckRobotCommandDatas As ObservableCollection(Of CommandData)
        Get
            Return Me._CheckRobotCommandDatas
        End Get
        Set(value As ObservableCollection(Of CommandData))
            Me.IsCheckRobotCommandDatasDirectCast = True
            Me._CheckRobotCommandDatas = value
            RaisePropertyChanged("CheckRobotCommandDatas")
            Me.IsCheckRobotCommandDatasDirectCast = False
        End Set
    End Property

    Private _CheckRobotCommandIndex As Integer
    <JsonIgnore>
    Public Property CheckRobotCommandIndex As Integer
        Get
            Return Me._CheckRobotCommandIndex
        End Get
        Set(value As Integer)
            If Me.IsCheckRobotCommandDatasDirectCast Then
                Me._CheckRobotCommandIndex = -1
            Else
                Me._CheckRobotCommandIndex = value
            End If
            RaisePropertyChanged("CheckRobotCommandIndex")
        End Set
    End Property

    Public Class CommandData
        Private _CommandText As String
        Public Property CommandText As String
            Get
                Return Me._CommandText
            End Get
            Set(value As String)
                Me._CommandText = value
            End Set
        End Property

        Private _ExecuteDate As Date
        Public Property ExecuteDate As Date
            Get
                Return Me._ExecuteDate
            End Get
            Set(value As Date)
                Me._ExecuteDate = value
            End Set
        End Property

        Private _ExecuteTimes As Integer
        Public Property ExecuteTimes As Integer
            Get
                Return Me._ExecuteTimes
            End Get
            Set(value As Integer)
                Me._ExecuteTimes = value
            End Set
        End Property
    End Class

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

    Private _AverageTime As Integer
    Public Property AverageTime As Integer
        Get
            Return Me._AverageTime
        End Get
        Set(value As Integer)
            Me._AverageTime = value
            RaisePropertyChanged("AverageTime")
        End Set
    End Property

    Private _RunAverageTime As Integer
    Public Property RunAverageTime As Integer
        Get
            Return Me._RunAverageTime
        End Get
        Set(value As Integer)
            Me._RunAverageTime = value
            RaisePropertyChanged("RunAverageTime")
        End Set
    End Property

    Private _CheckAverageTime As Integer
    Public Property CheckAverageTime As Integer
        Get
            Return Me._CheckAverageTime
        End Get
        Set(value As Integer)
            Me._CheckAverageTime = value
            RaisePropertyChanged("CheckAverageTime")
        End Set
    End Property

    Private _RunProcessTimeLogs As List(Of Integer)
    Public Property RunProcessTimeLogs As List(Of Integer)
        Get
            Return Me._RunProcessTimeLogs
        End Get
        Set(value As List(Of Integer))
            Me._RunProcessTimeLogs = value
        End Set
    End Property

    Private _CheckProcessTimeLogs As List(Of Integer)
    Public Property CheckProcessTimeLogs As List(Of Integer)
        Get
            Return Me._CheckProcessTimeLogs
        End Get
        Set(value As List(Of Integer))
            Me._CheckProcessTimeLogs = value
        End Set
    End Property

    Private _IsProjectExist As Boolean
    Private Property IsProjectExist As Boolean
        Get
            Return Me._IsProjectExist
        End Get
        Set(value As Boolean)
            Me._IsProjectExist = value
        End Set
    End Property

    Private _IsProcessEnded As Boolean
    Private Property IsProcessEnded As Boolean
        Get
            Return Me._IsProcessEnded
        End Get
        Set(value As Boolean)
            Me._IsProcessEnded = value
        End Set
    End Property

    Private _OverTimeFlag As Boolean
    Public Property OverTimeFlag As Boolean
        Get
            Return Me._OverTimeFlag
        End Get
        Set(value As Boolean)
            Me._OverTimeFlag = value
            RaisePropertyChanged("RunRobotCommandToolTip")
        End Set
    End Property

    Private _MyRobotReadMe As String
    <JsonIgnore>
    Public Property MyRobotReadMe As String
        Get
            If String.IsNullOrEmpty(Me._MyRobotReadMe) Then
                If ViewController.Data IsNot Nothing Then
                    If ViewController.Data.Project IsNot Nothing Then
                        Me._MyRobotReadMe = ViewController.Data.Project.MyRobotReadMe
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

    Private _ProjectRunner As ProjectRunnerModel
    Public Property ProjectRunner As ProjectRunnerModel
        Get
            If Me._ProjectRunner Is Nothing Then
                Me._ProjectRunner = New ProjectRunnerModel
            End If
            Return Me._ProjectRunner
        End Get
        Set(value As ProjectRunnerModel)
            Me._ProjectRunner = value
            RaisePropertyChanged("ProjectRunner")
        End Set
    End Property

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

    Private Sub RunRobotCommandExecute(ByVal parameter As Object)
        Call SetConsoleCommand(Me.RunRobotCommandDatas(0).CommandText)
        ViewController.ExecuteGuiCommandPath = True
    End Sub

    Private Function RunRobotCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsRunRobotCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    Private _RunRobotWithParametersCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property RunRobotWithParametersCommand As ICommand
        Get
            If Me._RunRobotWithParametersCommand Is Nothing Then
                Me._RunRobotWithParametersCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf RunRobotWithParametersCommandExecute,
                    .CanExecuteHandler = AddressOf RunRobotWithParametersCommandCanExecute
                }
            End If
            Return Me._RunRobotWithParametersCommand
        End Get
    End Property

    Private _IsRunRobotWithParametersCommandEnabled As Boolean
    <JsonIgnore>
    Public Property IsRunRobotWithParametersCommandEnabled As Boolean
        Get
            Return Me._IsRunRobotWithParametersCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsRunRobotWithParametersCommandEnabled = value
            RaisePropertyChanged("IsRunRobotWithParametersCommandEnabled")
            CType(RunRobotWithParametersCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub RunRobotWithParametersCommandExecute(ByVal parameter As Object)
        Call SetConsoleCommand(Me.RunRobotCommandText)
        ViewController.ExecuteGuiCommandPath = True
    End Sub

    Private Function RunRobotWithParametersCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsRunRobotWithParametersCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    Private _RunRobotWithNoParameterCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property RunRobotWithNoParameterCommand As ICommand
        Get
            If Me._RunRobotWithNoParameterCommand Is Nothing Then
                Me._RunRobotWithNoParameterCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf RunRobotWithNoParameterCommandExecute,
                    .CanExecuteHandler = AddressOf RunRobotWithNoParameterCommandCanExecute
                }
            End If
            Return Me._RunRobotWithNoParameterCommand
        End Get
    End Property

    Private _IsRunRobotWithNoParameterCommandEnabled As Boolean
    <JsonIgnore>
    Public Property IsRunRobotWithNoParameterCommandEnabled As Boolean
        Get
            Return Me._IsRunRobotWithNoParameterCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsRunRobotWithNoParameterCommandEnabled = value
            RaisePropertyChanged("IsRunRobotWithNoParameterCommandEnabled")
            CType(RunRobotWithNoParameterCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub RunRobotWithNoParameterCommandExecute(ByVal parameter As Object)
        Call SetConsoleCommand(RunnerViewModel.RUNCOMMAND)
        ViewController.ExecuteGuiCommandPath = True
    End Sub

    Private Function RunRobotWithNoParameterCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsRunRobotWithNoParameterCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

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

    Private Sub CheckRobotCommandExecute(ByVal parameter As Object)
        Call SetConsoleCommand(Me.CheckRobotCommandDatas(0).CommandText)
        ViewController.ExecuteGuiCommandPath = True
    End Sub

    Private Function CheckRobotCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsCheckRobotCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    Private _CheckRobotWithNoParameterCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property CheckRobotWithNoParameterCommand As ICommand
        Get
            If Me._CheckRobotWithNoParameterCommand Is Nothing Then
                Me._CheckRobotWithNoParameterCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf CheckRobotWithNoParameterCommandExecute,
                    .CanExecuteHandler = AddressOf CheckRobotWithNoParameterCommandCanExecute
                }
            End If
            Return Me._CheckRobotWithNoParameterCommand
        End Get
    End Property

    Private _IsCheckRobotWithNoParameterCommandEnabled As Boolean
    <JsonIgnore>
    Public Property IsCheckRobotWithNoParameterCommandEnabled As Boolean
        Get
            Return Me._IsCheckRobotWithNoParameterCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsCheckRobotWithNoParameterCommandEnabled = value
            RaisePropertyChanged("IsCheckRobotWithNoParameterCommandEnabled")
            CType(CheckRobotWithNoParameterCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub CheckRobotWithNoParameterCommandExecute(ByVal parameter As Object)
        Call SetConsoleCommand(RunnerViewModel.CHKCOMMAND)
        ViewController.ExecuteGuiCommandPath = True
    End Sub

    Private Function CheckRobotWithNoParameterCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsCheckRobotWithNoParameterCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    Private _CheckRobotWithParametersCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property CheckRobotWithParametersCommand As ICommand
        Get
            If Me._CheckRobotWithParametersCommand Is Nothing Then
                Me._CheckRobotWithParametersCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf CheckRobotWithParametersCommandExecute,
                    .CanExecuteHandler = AddressOf CheckRobotWithParametersCommandCanExecute
                }
            End If
            Return Me._CheckRobotWithParametersCommand
        End Get
    End Property

    Private _IsCheckRobotWithParametersCommandEnabled As Boolean
    <JsonIgnore>
    Public Property IsCheckRobotWithParametersCommandEnabled As Boolean
        Get
            Return Me._IsCheckRobotWithParametersCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsCheckRobotWithParametersCommandEnabled = value
            RaisePropertyChanged("IsCheckRobotWithParametersCommandEnabled")
            CType(CheckRobotWithParametersCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub CheckRobotWithParametersCommandExecute(ByVal parameter As Object)
        Call SetConsoleCommand(Me.CheckRobotCommandText)
        ViewController.ExecuteGuiCommandPath = True
    End Sub

    Private Function CheckRobotWithParametersCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsCheckRobotWithParametersCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    Private _SaveRunnerCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property SaveRunnerCommand As ICommand
        Get
            If Me._SaveRunnerCommand Is Nothing Then
                Me._SaveRunnerCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf SaveRunnerCommandExecute,
                    .CanExecuteHandler = AddressOf SaveRunnerCommandCanExecute
                }
            End If
            Return Me._SaveRunnerCommand
        End Get
    End Property

    Private _IsSaveRunnerCommandEnabled As Boolean
    <JsonIgnore>
    Public Property IsSaveRunnerCommandEnabled As Boolean
        Get
            Return Me._IsSaveRunnerCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsSaveRunnerCommandEnabled = value
            RaisePropertyChanged("IsSaveRunnerCommandEnabled")
            CType(SaveRunnerCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub SaveRunnerCommandExecute(ByVal parameter As Object)
        Call Me.Save(Me.RobotRunnerJsonFileName, Me)
        Call Me.ProjectRunner.Save(Me.ProjectRunnerJsonFileName, Me.ProjectRunner)
    End Sub

    Private Function SaveRunnerCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsSaveRunnerCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    Private _SaveMyRobotReadMeCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property SaveMyRobotReadMeCommand As ICommand
        Get
            If Me._SaveMyRobotReadMeCommand Is Nothing Then
                Me._SaveMyRobotReadMeCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf SaveMyRobotReadMeCommandExecute,
                    .CanExecuteHandler = AddressOf SaveMyRobotReadMeCommandCanExecute
                }
            End If
            Return Me._SaveMyRobotReadMeCommand
        End Get
    End Property

    Private _IsSaveMyRobotReadMeCommandEnabled As Boolean
    <JsonIgnore>
    Public Property IsSaveMyRobotReadMeCommandEnabled As Boolean
        Get
            Return Me._IsSaveMyRobotReadMeCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsSaveMyRobotReadMeCommandEnabled = value
            RaisePropertyChanged("IsSaveMyRobotReadMeCommandEnabled")
            CType(SaveMyRobotReadMeCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub SaveMyRobotReadMeCommandExecute(ByVal parameter As Object)
        Dim sw As New StreamWriter(ViewController.Data.Project.MyRobotReadMeFileName, False, Text.Encoding.GetEncoding(Rpa00.RpaModule.DEFUALTENCODING))
        sw.Write(Me.MyRobotReadMe)
        sw.Dispose()
        sw.Close()
    End Sub

    Private Function SaveMyRobotReadMeCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsSaveMyRobotReadMeCommandEnabled
    End Function

    Private Function CheckSaveMyRobotReadMeCommandEnabled() As Boolean
        If Me.IsProjectExist Then
            Return True
        End If
        Return False
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

    '---------------------------------------------------------------------------------------------'
    Private _CollapseMyRobotReadMeCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property CollapseMyRobotReadMeCommand As ICommand
        Get
            If Me._CollapseMyRobotReadMeCommand Is Nothing Then
                Me._CollapseMyRobotReadMeCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf CollapseMyRobotReadMeCommandExecute,
                    .CanExecuteHandler = AddressOf CollapseMyRobotReadMeCommandCanExecute
                }
            End If
            Return Me._CollapseMyRobotReadMeCommand
        End Get
    End Property

    Private _IsCollapseMyRobotReadMeCommandEnabled As Boolean
    <JsonIgnore>
    Public Property IsCollapseMyRobotReadMeCommandEnabled As Boolean
        Get
            Return Me._IsCollapseMyRobotReadMeCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsCollapseMyRobotReadMeCommandEnabled = value
            RaisePropertyChanged("IsCollapseMyRobotReadMeCommandEnabled")
            CType(CollapseMyRobotReadMeCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub CollapseMyRobotReadMeCommandExecute(ByVal parameter As Object)
        Me.MyRobotReadMeContentVisibility = Visibility.Collapsed
    End Sub

    Private Function CollapseMyRobotReadMeCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsCollapseMyRobotReadMeCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Sub CheckPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If e.PropertyName = "RunRobotCommandIndex" Then
            If sender.RunRobotCommandIndex >= 0 Then
                Call SetConsoleCommand(sender.RunRobotCommandDatas(sender.RunRobotCommandIndex).CommandText)
                Me.RunRobotCommandText = Me.RunRobotCommandDatas(sender.RunRobotCommandIndex).CommandText
                Me.CheckRobotCommandIndex = -1
            End If
        End If
        If e.PropertyName = "CheckRobotCommandIndex" Then
            If sender.CheckRobotCommandIndex >= 0 Then
                Call SetConsoleCommand(sender.CheckRobotCommandDatas(sender.CheckRobotCommandIndex).CommandText)
                Me.CheckRobotCommandText = Me.CheckRobotCommandDatas(sender.CheckRobotCommandIndex).CommandText
                Me.RunRobotCommandIndex = -1
            End If
        End If
    End Sub

    ' 廃止
    'Private Overloads Sub SetConsoleCommand(ByRef cmd As CommandData)
    '    ViewController.GuiCommandTextPath = cmd.CommandText
    'End Sub

    Private Overloads Sub SetConsoleCommand(ByVal cmdtxt As String)
        ViewController.GuiCommandTextPath = cmdtxt
    End Sub

    ' 検討中
    Private Sub CheckControllerPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If e.PropertyName = "GuiCommandExecuteStatusSender" Then
            Call CheckProcessTime(ViewController)
            Exit Sub
        End If
        If e.PropertyName = "GuiCommandTextPath" Then
            Me.GuiCommand = ViewController.GuiCommandTextPath
            Exit Sub
        End If
    End Sub

    Private Async Sub CheckProcessTime(ByVal sender As Object)
        If sender.GuiCommandExecuteStatusSender = 0 Then
            Me.IsProcessEnded = True
            Exit Sub
        End If
        If sender.GuiCommandExecuteStatusSender = 1 Then
            Me.IsProcessEnded = False
            Dim t As Task = Task.Run(
                Sub()
                    Call ProcessStart()
                End Sub
            )
            Await Task.WhenAll(t)
            Exit Sub
        End If
    End Sub

    Private Sub CalculateRunAverageTime()
        Dim totaltime As Integer = Me.RunProcessTimeLogs.Sum
        Me.RunAverageTime = (totaltime / Me.RunProcessTimeLogs.Count)
    End Sub

    Private Sub CalculateCheckAverageTime()
        Dim totaltime As Integer = Me.CheckProcessTimeLogs.Sum
        Me.CheckAverageTime = (totaltime / Me.CheckProcessTimeLogs.Count)
    End Sub

    Private Sub ProcessEnd()
        Dim cmds As String() = Me.GuiCommand.Split(" "c)
        If Me.ProcessTime > 0 Then
            If cmds(0) = RunnerViewModel.RUNCOMMAND Then
                If ViewController.Data.System.ReturnCode = 0 Then
                    Me.RunProcessTimeLogs = Rpa00.RpaModule.Push(Of Integer, List(Of Integer))(Me.ProcessTime, Me.RunProcessTimeLogs)
                    Call CalculateRunAverageTime()
                End If
                Me.ProjectRunner.ExecutionsOfYear.Last.Count += 1
            End If
            If cmds(0) = RunnerViewModel.CHKCOMMAND Then
                If ViewController.Data.System.ReturnCode = 0 Then
                    Me.CheckProcessTimeLogs = Rpa00.RpaModule.Push(Of Integer, List(Of Integer))(Me.ProcessTime, Me.CheckProcessTimeLogs)
                    Call CalculateCheckAverageTime()
                End If
                Me.ProjectRunner.ExecutionsOfYear.Last.Count += 1
            End If
            Me.ProcessTime = 0
        End If
    End Sub

    Private Sub ProcessStart()
        Me.OverTimeFlag = False
        Dim cmds As String() = Me.GuiCommand.Split(" "c)
        If cmds(0) = RunnerViewModel.RUNCOMMAND Then
            Me.AverageTime = Me.RunAverageTime
        End If
        If cmds(0) = RunnerViewModel.CHKCOMMAND Then
            Me.AverageTime = Me.CheckAverageTime
        End If

        Do
            Threading.Thread.Sleep(1000)
            Me.ProcessTime += 1
            If Me.ProcessTime > Me.AverageTime Then
                Me.OverTimeFlag = True
                Exit Do
            End If
            If Me.IsProcessEnded Then
                Exit Do
            End If
        Loop Until False

        Do
            If Me.IsProcessEnded Then
                Exit Do
            End If
            Threading.Thread.Sleep(1000)
            Me.ProcessTime += 1
        Loop Until False

        Call ProcessEnd()
    End Sub

    Private Sub UpdateRunRobotCommandDatas(ByVal cmdstr As String)
        Dim times As Integer = 0
        For Each cmddata In Me.RunRobotCommandDatas
            If cmddata.CommandText = cmdstr Then
                times = cmddata.ExecuteTimes
                Me.RunRobotCommandDatas.Remove(cmddata)
                Exit For
            End If
        Next
        Dim newcmd As CommandData = New CommandData With {.CommandText = cmdstr, .ExecuteDate = DateTime.Now, .ExecuteTimes = (times + 1)}
        Me.RunRobotCommandDatas = Rpa00.RpaModule.Push(Of CommandData, ObservableCollection(Of CommandData))(newcmd, Me.RunRobotCommandDatas)
        Me.RunRobotCommandToolTip = vbNullString
    End Sub

    Private Sub UpdateCheckRobotCommandDatas(ByVal cmdstr As String)
        Dim times As Integer = 0
        For Each cmddata In Me.CheckRobotCommandDatas
            If cmddata.CommandText = cmdstr Then
                times = cmddata.ExecuteTimes
                Me.CheckRobotCommandDatas.Remove(cmddata)
                Exit For
            End If
        Next
        Dim newcmd As CommandData = New CommandData With {.CommandText = cmdstr, .ExecuteDate = DateTime.Now, .ExecuteTimes = (times + 1)}
        Me.CheckRobotCommandDatas = Rpa00.RpaModule.Push(Of CommandData, ObservableCollection(Of CommandData))(newcmd, Me.CheckRobotCommandDatas)
        Me.CheckRobotCommandToolTip = vbNullString
    End Sub

    ' ConsoleOutputViewModel.GuiCommandTextsPathにアイテムが追加された時、こちらの履歴にも追加する
    ' ConsoleOutputViewModel.GuiCommandTextsPathにアイテムが追加されるのは、実行時となる
    Private Sub CheckCollectionChanged(ByVal sender As Object, ByVal e As NotifyCollectionChangedEventArgs)
        If e.Action = NotifyCollectionChangedAction.Add Then
            'Dim cmdstr As String = e.NewItems(e.NewStartingIndex)
            Dim cmdstr As String = ViewController.GuiCommandTextsPath(e.NewStartingIndex)
            Dim cmds As String() = cmdstr.Split(" "c)
            If cmds(0) = RunnerViewModel.RUNCOMMAND Then
                UpdateRunRobotCommandDatas(cmdstr)
                Exit Sub
            End If
            If cmds(0) = RunnerViewModel.CHKCOMMAND Then
                UpdateCheckRobotCommandDatas(cmdstr)
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Initialize()
        If ViewController IsNot Nothing Then
            ViewController.ExecutableGuiCommands = New List(Of String)(New String() {RunnerViewModel.RUNCOMMAND, RunnerViewModel.CHKCOMMAND})
            If ViewController.Data IsNot Nothing Then
                If ViewController.Data.Project IsNot Nothing Then
                    Me.IsProjectExist = True
                    Dim vm As RunnerViewModel = Me.Load(Me.RobotRunnerJsonFileName)
                    If vm Is Nothing Then
                        Me.RunRobotCommandDatas = New ObservableCollection(Of CommandData)
                        Me.RunRobotCommandDatas.Add(New CommandData With {.CommandText = RunnerViewModel.RUNCOMMAND, .ExecuteDate = DateTime.Now})
                        Me.CheckRobotCommandDatas = New ObservableCollection(Of CommandData)
                        Me.CheckRobotCommandDatas.Add(New CommandData With {.CommandText = RunnerViewModel.CHKCOMMAND, .ExecuteDate = DateTime.Now})

                        Me.RunProcessTimeLogs = New List(Of Integer)
                        Me.CheckProcessTimeLogs = New List(Of Integer)

                        Me.AverageTime = 60
                        Me.RunAverageTime = 60
                        Me.CheckAverageTime = 60
                    Else
                        Me.RunRobotCommandDatas = vm.RunRobotCommandDatas
                        Me.CheckRobotCommandDatas = vm.CheckRobotCommandDatas

                        Me.RunProcessTimeLogs = vm.RunProcessTimeLogs
                        Me.CheckProcessTimeLogs = vm.CheckProcessTimeLogs

                        Me.AverageTime = vm.AverageTime
                        Me.RunAverageTime = vm.RunAverageTime
                        Me.CheckAverageTime = vm.CheckAverageTime
                    End If

                    Me.ProjectRunner = Me.ProjectRunner.Load(Me.ProjectRunnerJsonFileName)
                    Call Me.ProjectRunner.GoExecutionsOfYearFoward()

                    Me.IsReloadMyRobotReadMeCommandEnabled = True
                    Me.IsSaveRunnerCommandEnabled = True
                    Me.CheckRobotCommandIndex = -1
                    Me.RunRobotCommandIndex = -1
                    Me.RunRobotCommandText = Me.RunRobotCommandDatas(0).CommandText
                    Me.CheckRobotCommandText = Me.CheckRobotCommandDatas(0).CommandText
                    AddHandler Me.PropertyChanged, AddressOf CheckPropertyChanged
                    AddHandler ViewController.PropertyChanged, AddressOf CheckControllerPropertyChanged
                    AddHandler ViewController.GuiCommandTextsPath.CollectionChanged, AddressOf CheckCollectionChanged
                End If
            End If
        End If
        Me.IsSaveMyRobotReadMeCommandEnabled = CheckSaveMyRobotReadMeCommandEnabled()
        Me.IsCollapseMyRobotReadMeCommandEnabled = True
        Me.IsRunRobotCommandEnabled = True
        Me.IsCheckRobotCommandEnabled = True
        Me.IsRunRobotWithParametersCommandEnabled = True
        Me.IsCheckRobotWithParametersCommandEnabled = True
        Me.IsRunRobotWithNoParameterCommandEnabled = True
        Me.IsCheckRobotWithNoParameterCommandEnabled = True
    End Sub

    Sub New()
    End Sub
End Class

Public Class ProjectRunnerModel : Inherits ViewModelBase(Of ProjectRunnerModel)
    Public Class ExecutionOfDay : Implements INotifyPropertyChanged

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
        Protected Sub RaisePropertyChanged(ByVal PropertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
        End Sub

        Private _ExecuteDate As Date
        Public Property ExecuteDate As Date
            Get
                Return Me._ExecuteDate
            End Get
            Set(value As Date)
                Me._ExecuteDate = value
                RaisePropertyChanged("ExecuteDate")
            End Set
        End Property

        Private _Count As Integer
        Public Property Count As Integer
            Get
                Return Me._Count
            End Get
            Set(value As Integer)
                Me._Count = value
                RaisePropertyChanged("Count")
            End Set
        End Property
    End Class

    Private _ExecutionsOfYear As ObservableCollection(Of ExecutionOfDay)
    Public Property ExecutionsOfYear As ObservableCollection(Of ExecutionOfDay)
        Get
            If Me._ExecutionsOfYear Is Nothing Then
                Me._ExecutionsOfYear = New ObservableCollection(Of ExecutionOfDay)
            End If
            Return Me._ExecutionsOfYear
        End Get
        Set(value As ObservableCollection(Of ExecutionOfDay))
            Me._ExecutionsOfYear = value
            RaisePropertyChanged("ExecutionsOfYear")
        End Set
    End Property

    Public Sub GoExecutionsOfYearFoward()
        If Me.ExecutionsOfYear.Count = 0 Then
            Call InitializeExecutionsOfYear()
        End If

        Dim latest As Date = Me.ExecutionsOfYear.Last.ExecuteDate
        Dim term As Integer = (DateTime.Today - latest).TotalDays

        If term >= 366 Then
            Call InitializeExecutionsOfYear()
            term = 0
        End If

        Dim i As Integer = term
        Do
            If i <= 0 Then
                Exit Do
            End If
            Me.ExecutionsOfYear.RemoveAt(0)
            Me.ExecutionsOfYear.Add(New ExecutionOfDay With {.ExecuteDate = DateTime.Today.AddDays(-i)})
            i -= 1
        Loop Until False
    End Sub

    Private Sub InitializeExecutionsOfYear()
        Me.ExecutionsOfYear = New ObservableCollection(Of ExecutionOfDay)
        Dim today As Date = DateTime.Today
        For i As Integer = 366 To 0 Step -1
            Me.ExecutionsOfYear.Add(New ExecutionOfDay With {.ExecuteDate = today.AddDays(-i)})
        Next
    End Sub
End Class
