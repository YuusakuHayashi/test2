Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Public Class ServerModel
    Inherits BaseModel
    Implements TreeViewInterface

    Private _Name As String
    Public Property Name As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
            'RaisePropertyChanged("Name")
        End Set
    End Property

    Private _DataBases As List(Of DataBaseModel)
    Public Property DataBases As List(Of DataBaseModel)
        Get
            Return Me._DataBases
        End Get
        Set(value As List(Of DataBaseModel))
            Me._DataBases = value
            RaisePropertyChanged("DataBases")
        End Set
    End Property

    Private _IsChecked As Boolean
    Public Property IsChecked As Boolean Implements TreeViewInterface.IsChecked
        Get
            Return Me._IsChecked
        End Get
        Set(value As Boolean)
            Me._IsChecked = value
            RaisePropertyChanged("_IsChecked")
            Call CheckingChildren(Of DataBaseModel)(DataBases, value)
        End Set
    End Property


    Private _IsEnabled As Boolean
    Public Property IsEnabled As Boolean Implements TreeViewInterface.IsEnabled
        Get
            Return Me._IsEnabled
        End Get
        Set(value As Boolean)
            Me._IsEnabled = value
        End Set
    End Property

    Sub New()
        Me.IsChecked = True
        Me.IsEnabled = True
    End Sub
End Class

