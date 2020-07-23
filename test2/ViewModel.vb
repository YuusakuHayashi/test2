Imports System.ComponentModel
Imports System.Collections.ObjectModel

Public Class ViewModel : Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Protected Sub RaisePropertyChanged(ByVal propertyName As String)
        RaiseEvent PropertyChanged(Me,
                                    New PropertyChangedEventArgs(propertyName))
    End Sub


    'Private _Name As String
    'Public Property Name As String
    '    Get
    '        Return _Name
    '    End Get
    '    Set(value As String)
    '        _Name = value
    '    End Set
    'End Property


    'Private _PluginEntries As List(Of Plugin)
    'Public Property PluginEntries As List(Of Plugin)
    '    Get
    '        Return _PluginEntries
    '    End Get
    '    Set(value As List(Of Plugin))
    '        _PluginEntries = value
    '    End Set
    'End Property

    Private _vm As List(Of Model)

    Public Property vm As List(Of Model)
        Get
            Return _vm
        End Get
        Set(value As List(Of Model))
            _vm = value
        End Set
    End Property

    Sub New()
        Me.vm = New List(Of Model)
        Me.vm.Add(New Model With {
            .ServerName = "AServer",
            .Child = New List(Of Model) From {
                New Model With {.ServerName = "BServer", .Child = New List(Of Model) From {
                    New Model With {.ServerName = "CServer", .Child = New List(Of Model) From {
                        New Model With {.ServerName = "DServer"}
                    }}
                }},
                New Model With {.ServerName = "B2Server", .Child = New List(Of Model) From {
                    New Model With {.ServerName = "C2Server", .Child = New List(Of Model) From {
                        New Model With {.ServerName = "D2Server"}
                    }}
                }}
            }
        })
    End Sub
End Class
