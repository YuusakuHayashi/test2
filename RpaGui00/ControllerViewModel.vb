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
                Me._ExplorerTabs = New ObservableCollection(Of TabItemViewModel)
                Me._ExplorerTabs.Add(New TabItemViewModel With {.Header = "メニュー", .Content = New MenuExplorerViewModel With {.ViewController = Me}})
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
                Me._OutputTabs = New ObservableCollection(Of TabItemViewModel)
                Me._OutputTabs.Add(New TabItemViewModel With {.Header = "コンソール", .Content = New ConsoleOutputViewModel With {.ViewController = Me}})
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

    Sub New()
        Me.ExplorerTabIndex = 0
        Me.OutputTabIndex = 0
    End Sub
End Class
