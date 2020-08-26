Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Public Class ServerModel : Inherits BaseModel
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

    Private _DataBases As ObservableCollection(Of DataBaseModel)
    Public Property DataBases As ObservableCollection(Of DataBaseModel)
        Get
            Return Me._DataBases
End Get
Set(value As ObservableCollection(Of DataBaseModel))
            Me._DataBases = value
            RaisePropertyChanged("DataBases")
        End Set
    End Property

    'Private _IsSelected As Boolean
    'Public Property IsSelected As Boolean
    '    Get
    '        Return Me._IsSelected
    '    End Get
    '    Set(value As Boolean)
    '        Me._IsSelected = value
    '        RaisePropertyChanged("_IsSelected")
    '    End Set
    'End Property

    Private _IsChecked As Boolean
    Public Property IsChecked As Boolean
        Get
            Return Me._IsChecked
        End Get
        Set(value As Boolean)
            Me._IsChecked = value
            RaisePropertyChanged("_IsChecked")
        End Set
    End Property


    Sub New()
    End Sub
End Class

