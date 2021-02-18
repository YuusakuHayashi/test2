﻿Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Newtonsoft.Json

Public Class ControllerViewModel : Inherits ViewModelBase(Of ControllerViewModel)
    Private ReadOnly Property ContentErrorImageFileName As String
        Get
            Return $"{Controller.SystemDllDirectory}\contenterror.png"
        End Get
    End Property
    Private _ContentErrorImage As BitmapImage
    Public ReadOnly Property ContentErrorImage As BitmapImage
        Get
            If Me._ContentErrorImage Is Nothing Then
                Me._ContentErrorImage = RpaGuiModule.CreateIcon(Me.ContentErrorImageFileName)
            End If
            Return Me._ContentErrorImage
        End Get
    End Property

    Private _CurrentProjectName As String
    Public Property CurrentProjectName As String
        Get
            Return Me._CurrentProjectName
        End Get
        Set(value As String)
            Me._CurrentProjectName = value
            Call RaisePropertyChanged("CurrentProjectName")
        End Set
    End Property

    Private _CurrentRobotName As String
    Public Property CurrentRobotName As String
        Get
            Return Me._CurrentRobotName
        End Get
        Set(value As String)
            Me._CurrentRobotName = value
            Call RaisePropertyChanged("CurrentRobotName")
        End Set
    End Property

    Private _Data As Rpa00.RpaDataWrapper
    <JsonIgnore>
    Public Property Data As Rpa00.RpaDataWrapper
        Get
            Return Me._Data
        End Get
        Set(value As Rpa00.RpaDataWrapper)
            Me._Data = value
            Call RaisePropertyChanged("Data")
        End Set
    End Property

    Private _StatusCode As Integer
    Public Property StatusCode As Integer
        Get
            Return Me._StatusCode
        End Get
        Set(value As Integer)
            Me._StatusCode = value
            Call RaisePropertyChanged("StatusCode")
        End Set
    End Property

    Private _ExplorerTabIndex As Integer
    Public Property ExplorerTabIndex As Integer
        Get
            Return Me._ExplorerTabIndex
        End Get
        Set(value As Integer)
            Me._ExplorerTabIndex = value
        End Set
    End Property

    Private _ExplorerTabs As ObservableCollection(Of TabItemViewModel)
    Public ReadOnly Property ExplorerTabs As ObservableCollection(Of TabItemViewModel)
        Get
            If Me._ExplorerTabs Is Nothing Then
                Dim [obj] As Object
                Dim [tab] As TabItemViewModel
                Me._ExplorerTabs = New ObservableCollection(Of TabItemViewModel)

                [obj] = New MenuExplorerViewModel With {.ViewController = Me}
                [obj].Initialize()
                [tab] = New TabItemViewModel With {.Header = "メニュー", .Content = obj}
                Me._ExplorerTabs.Add([tab])
            End If
            Return Me._ExplorerTabs
        End Get
    End Property

    Private _OutputTabIndex As Integer
    Public Property OutputTabIndex As Integer
        Get
            Return Me._OutputTabIndex
        End Get
        Set(value As Integer)
            Me._OutputTabIndex = value
        End Set
    End Property

    Private _OutputTabs As ObservableCollection(Of TabItemViewModel)
    Public ReadOnly Property OutputTabs As ObservableCollection(Of TabItemViewModel)
        Get
            If Me._OutputTabs Is Nothing Then
                Dim [obj] As Object
                Dim [tab] As TabItemViewModel
                Me._OutputTabs = New ObservableCollection(Of TabItemViewModel)

                [obj] = New ConsoleOutputViewModel With {.ViewController = Me}
                [obj].Initialize()
                [tab] = New TabItemViewModel With {.Header = "コンソール", .Content = obj}
                Me._OutputTabs.Add([tab])

                [obj] = New CommandLogsOutputViewModel With {.ViewController = Me}
                [obj].Initialize()
                [tab] = New TabItemViewModel With {.Header = "コマンドログ", .Content = obj}
                Me._OutputTabs.Add([tab])
            End If
            Return Me._OutputTabs
        End Get
    End Property

    Private _MainContent As Object
    <JsonIgnore>
    Public Property MainContent As Object
        Get
            If Me._MainContent Is Nothing Then
                Me._MainContent = New NoContentViewModel
            End If
            Return Me._MainContent
        End Get
        Set(value As Object)
            Me._MainContent = value
            Call RaisePropertyChanged("MainContent")
        End Set
    End Property

    ' Guiコマンドの制御
    Private _ExecutableGuiCommands As List(Of String)
    Public Property ExecutableGuiCommands As List(Of String)
        Get
            If Me._ExecutableGuiCommands Is Nothing Then
                Me._ExecutableGuiCommands = New List(Of String)
            End If
            Return Me._ExecutableGuiCommands
        End Get
        Set(value As List(Of String))
            Me._ExecutableGuiCommands = value
        End Set
    End Property

    ' ConsoleOutputViewModel の GuiCommandText を変更するためのプロパティ
    ' ConsoleOutputViewModel は、このプロパティの変更を監視しています
    Private _GuiCommandChangeSender As String
    Public Property GuiCommandChangeSender As String
        Get
            Return Me._GuiCommandChangeSender
        End Get
        Set(value As String)
            Me._GuiCommandChangeSender = value
            RaisePropertyChanged("GuiCommandChangeSender")
        End Set
    End Property

    Private _GuiCommandTextPath As String
    Public Property GuiCommandTextPath As String
        Get
            Return Me._GuiCommandTextPath
        End Get
        Set(value As String)
            Me._GuiCommandTextPath = value
            RaisePropertyChanged("GuiCommandTextPath")
        End Set
    End Property

    Private _GuiCommandTextsPath As ObservableCollection(Of String)
    Public Property GuiCommandTextsPath As ObservableCollection(Of String)
        Get
            If Me._GuiCommandTextsPath Is Nothing Then
                Me._GuiCommandTextsPath = New ObservableCollection(Of String)
            End If
            Return Me._GuiCommandTextsPath
        End Get
        Set(value As ObservableCollection(Of String))
            Me._GuiCommandTextsPath = value
            RaisePropertyChanged("GuiCommandTextsPath")
        End Set
    End Property

    ' コマンドが実行されたかどうかを
    Private _ExecuteGuiCommandExecutedSender As Boolean
    Public Property ExecuteGuiCommandExecutedSender As Boolean
        Get
            Return Me._ExecuteGuiCommandExecutedSender
        End Get
        Set(value As Boolean)
            Me._ExecuteGuiCommandExecutedSender = value
        End Set
    End Property

    ' コマンドが実行中かどうか
    Private _GuiCommandExecuteStatusSender As Integer
    Public Property GuiCommandExecuteStatusSender As Integer
        Get
            Return Me._GuiCommandExecuteStatusSender
        End Get
        Set(value As Integer)
            Me._GuiCommandExecuteStatusSender = value
            RaisePropertyChanged("GuiCommandExecuteStatusSender")
        End Set
    End Property

    'Private _CommandLogsPath As Object
    'Public Property CommandLogsPath As Object
    '    Get
    '        Return Me._CommandLogsPath
    '    End Get
    '    Set(value As Object)
    '        Me._CommandLogsPath = value
    '        RaisePropertyChanged("CommandLogsPath")
    '    End Set
    'End Property

    ' ConsoleOutputクラスがExecuteGuiCommand起動要求を監視するためのパス
    Private _ExecuteGuiCommandPath As Boolean
    Public Property ExecuteGuiCommandPath As Boolean
        Get
            Return Me._ExecuteGuiCommandPath
        End Get
        Set(value As Boolean)
            Me._ExecuteGuiCommandPath = value
            RaisePropertyChanged("ExecuteGuiCommandPath")
        End Set
    End Property
    '---------------------------------------------------------------------------------------------'

    Private _ExitCommand As ICommand
    Public ReadOnly Property ExitCommand As ICommand
        Get
            If Me._ExitCommand Is Nothing Then
                Me._ExitCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf ExitCommandExecute,
                    .CanExecuteHandler = AddressOf ExitCommandCanExecute
                }
            End If
            Return Me._ExitCommand
        End Get
    End Property

    Private _IsExitCommandEnabled As Boolean
    Public Property IsExitCommandEnabled As Boolean
        Get
            Return Me._IsExitCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsExitCommandEnabled = value
            RaisePropertyChanged("IsExitCommandEnabled")
            CType(ExitCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Async Sub ExitCommandExecute(ByVal parameter As Object)
        Dim t As Task = Task.Run(
            Sub()
                Call _ExitCommandExecute(parameter)
            End Sub
        )
        Await Task.WhenAll(t)
        Application.Current.Shutdown()
    End Sub

    Private Sub _ExitCommandExecute(ByVal parameter As Object)
        Data.System.ExecuteMode = Rpa00.RpaSystem.ExecuteModeNumber.RpaGui
        Data.System.GuiCommandText = $"exit"
        Call Data.System.Main(Data)
    End Sub

    Private Function ExitCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsExitCommandEnabled
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Sub CheckDataChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If e.PropertyName = "CurrentProject" Then
            Call SetCurrentProjectName()
        End If
    End Sub

    Private Sub SetCurrentProjectName()
        If Data.Initializer.CurrentProject Is Nothing Then
            Me.CurrentProjectName = "No Project"
        End If
        If String.IsNullOrEmpty(Data.Initializer.CurrentProject.Name) Then
            Me.CurrentProjectName = "No Project"
        Else
            Me.CurrentProjectName = Data.Initializer.CurrentProject.Name
        End If
    End Sub

    Private Sub CheckProjectChanged(ByVal sender As Object, ByVal e As Rpa00.RpaEvent.ProjectChangedEventArgs)
        If e.PropertyName = "RobotName" Then
            Call SetCurrentRobotName()
        End If
    End Sub

    Private Sub SetCurrentRobotName()
        If String.IsNullOrEmpty(Data.Project.RobotAlias) Then
            Me.CurrentProjectName = "No Robot"
        Else
            Me.CurrentRobotName = Data.Project.RobotAlias
        End If
    End Sub

    Public Overrides Sub Initialize()
        Data.System.ReturnCode = -1
        Call SetCurrentProjectName()
        Call SetCurrentRobotName()
        AddHandler Data.Initializer.PropertyChanged, AddressOf CheckDataChanged
        AddHandler Rpa00.RpaEvent.Instance.ProjectChanged, AddressOf CheckProjectChanged
    End Sub

    Sub New()
        Me.IsExitCommandEnabled = True
        Me.ExplorerTabIndex = 0
        Me.OutputTabIndex = 0
    End Sub
End Class
