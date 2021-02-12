Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Newtonsoft.Json

Public Class ControllerViewModel : Inherits ViewModelBase(Of ControllerViewModel)
    ' !!! RpaDataWrapper型のdllオブジェクトがキャストされる
    ' 別プロジェクトのせいか、RpaDataWrapper型としては定義できなかった
    ' (するとキャスト時エラーになる)
    Private _Data As Object
    <JsonIgnore>
    Public Property Data As Object
        Get
            Return Me._Data
        End Get
        Set(value As Object)
            Me._Data = value
            Call RaisePropertyChanged("Data")
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

                'Me._ExplorerTabs.Add(New TabItemViewModel With {.Header = "メニュー", .Content = New MenuExplorerViewModel With {.ViewController = Me}})
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

                'Me._OutputTabs.Add(New TabViewModel With {.Header = "コンソール", .Content = New ConsoleOutputViewModel With {.ViewController = Me}})
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

    'Private Sub CheckPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
    '    If e.PropertyName = "ExplorerTabIndex" Then
    '        Me.ExplorerTabs(Me.OutputTabIndex).Content.ViewController = Me
    '        Me.ExplorerTabs(Me.OutputTabIndex).Content.Initialize()
    '    End If
    '    If e.PropertyName = "OutputTabIndex" Then
    '        Me.OutputTabs(Me.OutputTabIndex).Content.ViewController = Me
    '        Me.OutputTabs(Me.OutputTabIndex).Content.Initialize()
    '    End If
    'End Sub

    Public Overrides Sub Initialize()
        Data.System.ReturnCode = -1
    End Sub

    Sub New()
        Me.ExplorerTabIndex = 0
        Me.OutputTabIndex = 0
        'AddHandler Me.PropertyChanged, AddressOf CheckPropertyChanged
    End Sub
End Class
