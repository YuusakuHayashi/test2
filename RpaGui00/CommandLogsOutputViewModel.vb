Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel

Public Class CommandLogsOutputViewModel : Inherits ControllerViewModelBase(Of CommandLogsOutputViewModel)
    Private _Logs As ObservableCollection(Of Rpa00.RpaCommand)
    Public Property Logs As ObservableCollection(Of Rpa00.RpaCommand)
        Get
            If Me._Logs Is Nothing Then
                Me._Logs = New ObservableCollection(Of Rpa00.RpaCommand)
            End If
            Return Me._Logs
        End Get
        Set(value As ObservableCollection(Of Rpa00.RpaCommand))
            Me._Logs = value
            RaisePropertyChanged("Logs")
        End Set
    End Property

    Private Sub CommandLogsChanged(ByVal sender As ObservableCollection(Of Rpa00.RpaCommand), ByVal e As NotifyCollectionChangedEventArgs)
        If e.Action = NotifyCollectionChangedAction.Add Then
            Me.Logs.Add(e.NewItems.SyncRoot(e.NewStartingIndex))
        End If
    End Sub

    Public Overrides Sub Initialize()
        AddHandler ViewController.Data.System.CommandLogs.CollectionChanged, AddressOf CommandLogsChanged
    End Sub
End Class
