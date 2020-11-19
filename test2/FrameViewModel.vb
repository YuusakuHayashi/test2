Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class FrameViewModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Public Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub

    Private _MenuViewHeight As Double
    Public Property MenuViewHeight As Double
        Get
            Return Me._MenuViewHeight
        End Get
        Set(value As Double)
            Me._MenuViewHeight = value
        End Set
    End Property

    Private _MenuContent As Object
    Public Property MenuContent As Object
        Get
            Return Me._MenuContent
        End Get
        Set(value As Object)
            Me._MenuContent = value
        End Set
    End Property

    Private _MainContent As Object
    Public Property MainContent As Object
        Get
            Return Me._MainContent
        End Get
        Set(value As Object)
            Me._MainContent = value
        End Set
    End Property

    Private _LeftExplorerContent As Object
    Public Property LeftExplorerContent As Object
        Get
            Return Me._LeftExplorerContent
        End Get
        Set(value As Object)
            Me._LeftExplorerContent = value
        End Set
    End Property

    Private _HistoryContent As Object
    Public Property HistoryContent As Object
        Get
            Return Me._HistoryContent
        End Get
        Set(value As Object)
            Me._HistoryContent = value
        End Set
    End Property

    Private _RightExplorerContent As Object
    Public Property RightExplorerContent As Object
        Get
            Return Me._RightExplorerContent
        End Get
        Set(value As Object)
            Me._RightExplorerContent = value
        End Set
    End Property
End Class
