﻿Imports System.ComponentModel
Imports System.Collections.Specialized
Public Class ViewModelBase : Implements INotifyPropertyChanged, INotifyCollectionChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Event CollectionChanged As NotifyCollectionChangedEventHandler Implements INotifyCollectionChanged.CollectionChanged

    Protected Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub

    Protected Sub RaiseCollectionChanged(ByVal PropertyName As String)
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(PropertyName))
    End Sub
End Class
