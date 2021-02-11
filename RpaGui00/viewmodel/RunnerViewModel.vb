Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Newtonsoft.Json

Public Class RunnerViewModel : Inherits ControllerViewModelBase(Of RunnerViewModel)
    Private ReadOnly Property RunnerJsonFileName As String
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

    Private _IsProcessEnded As Boolean
    Private Property IsProcessEnded As Boolean
        Get
            Return Me._IsProcessEnded
        End Get
        Set(value As Boolean)
            Me._IsProcessEnded = value
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
        Call Me.Save(Me.RunnerJsonFileName, Me)
    End Sub

    Private Function SaveRunnerCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsSaveRunnerCommandEnabled
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

    Private Sub CheckPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If e.PropertyName = "RunRobotCommandIndex" Then
            If sender.RunRobotCommandIndex >= 0 Then
                Call SetConsoleCommand(sender.RunRobotCommandDatas(sender.RunRobotCommandIndex))
                Me.CheckRobotCommandIndex = -1
            End If
        End If
        If e.PropertyName = "CheckRobotCommandIndex" Then
            If sender.CheckRobotCommandIndex >= 0 Then
                Call SetConsoleCommand(sender.CheckRobotCommandDatas(sender.CheckRobotCommandIndex))
                Me.RunRobotCommandIndex = -1
            End If
        End If
    End Sub

    Private Sub SetConsoleCommand(ByRef cmd As CommandData)
        ViewController.GuiCommandTextPath = cmd.CommandText
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
            If cmds(0) = "runrobot" Then
                Me.RunProcessTimeLogs = Rpa00.RpaModule.Push(Of Integer, List(Of Integer))(Me.ProcessTime, Me.RunProcessTimeLogs)
                Call CalculateRunAverageTime()
            End If
            If cmds(0) = "checkrobot" Then
                Me.CheckProcessTimeLogs = Rpa00.RpaModule.Push(Of Integer, List(Of Integer))(Me.ProcessTime, Me.CheckProcessTimeLogs)
                Call CalculateCheckAverageTime()
            End If
            Me.ProcessTime = 0
        End If
    End Sub

    Private Sub ProcessStart()
        Dim cmds As String() = Me.GuiCommand.Split(" "c)
        If cmds(0) = "runrobot" Then
            Me.AverageTime = Me.RunAverageTime
        End If
        If cmds(0) = "checkrobot" Then
            Me.AverageTime = Me.CheckAverageTime
        End If

        Do
            Threading.Thread.Sleep(1000)
            Me.ProcessTime += 1
        Loop Until Me.IsProcessEnded

        Call ProcessEnd()
    End Sub

    Private Sub UpdateRunRobotCommandDatas(ByVal cmdstr As String)
        For Each cmddata In Me.RunRobotCommandDatas
            If cmddata.CommandText = cmdstr Then
                cmddata.ExecuteDate = DateTime.Now
                Exit Sub
            End If
        Next
        ' 新しいコマンドの場合
        Dim newcmd As CommandData = New CommandData With {.CommandText = cmdstr, .ExecuteDate = DateTime.Now}
        Me.RunRobotCommandDatas = Rpa00.RpaModule.Push(Of CommandData, ObservableCollection(Of CommandData))(newcmd, Me.RunRobotCommandDatas)
    End Sub

    Private Sub UpdateCheckRobotCommandDatas(ByVal cmdstr As String)
        For Each cmddata In Me.CheckRobotCommandDatas
            If cmddata.CommandText = cmdstr Then
                cmddata.ExecuteDate = DateTime.Now
                Exit Sub
            End If
        Next
        ' 新しいコマンドの場合
        Dim newcmd As CommandData = New CommandData With {.CommandText = cmdstr, .ExecuteDate = DateTime.Now}
        Me.CheckRobotCommandDatas = Rpa00.RpaModule.Push(Of CommandData, ObservableCollection(Of CommandData))(newcmd, Me.CheckRobotCommandDatas)
    End Sub

    Private Sub CheckCollectionChanged(ByVal sender As Object, ByVal e As NotifyCollectionChangedEventArgs)
        If e.Action = NotifyCollectionChangedAction.Add Then
            'Dim cmdstr As String = e.NewItems(e.NewStartingIndex)
            Dim cmdstr As String = ViewController.GuiCommandTextsPath(e.NewStartingIndex)
            Dim cmds As String() = cmdstr.Split(" "c)
            If cmds(0) = "runrobot" Then
                UpdateRunRobotCommandDatas(cmdstr)
                Exit Sub
            End If
            If cmds(0) = "checkrobot" Then
                UpdateCheckRobotCommandDatas(cmdstr)
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Initialize()
        If ViewController IsNot Nothing Then
            ViewController.ExecutableGuiCommands = New List(Of String)(New String() {"runrobot", "checkrobot"})
            If ViewController.Data IsNot Nothing Then
                If ViewController.Data.Project IsNot Nothing Then
                    Dim vm As RunnerViewModel = Me.Load(Me.RunnerJsonFileName)
                    If vm Is Nothing Then
                        Me.RunRobotCommandDatas = New ObservableCollection(Of CommandData)
                        Me.RunRobotCommandDatas.Add(New CommandData With {.CommandText = "runrobot", .ExecuteDate = DateTime.Now})
                        Me.CheckRobotCommandDatas = New ObservableCollection(Of CommandData)
                        Me.CheckRobotCommandDatas.Add(New CommandData With {.CommandText = "checkrobot", .ExecuteDate = DateTime.Now})

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
                    Me.IsReloadMyRobotReadMeCommandEnabled = True
                    Me.IsSaveRunnerCommandEnabled = True
                    Me.CheckRobotCommandIndex = -1
                    Me.RunRobotCommandIndex = -1
                    AddHandler Me.PropertyChanged, AddressOf CheckPropertyChanged
                    AddHandler ViewController.PropertyChanged, AddressOf CheckControllerPropertyChanged
                    AddHandler ViewController.GuiCommandTextsPath.CollectionChanged, AddressOf CheckCollectionChanged
                End If
            End If
        End If
    End Sub

    Sub New()
    End Sub
End Class
