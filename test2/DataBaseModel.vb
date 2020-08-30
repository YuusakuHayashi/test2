Imports System.Collections.ObjectModel

Public Class DataBaseModel
    Inherits BaseModel
    Implements TreeViewInterface


    Private _Name As String
    Public Property Name As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
            RaisePropertyChanged("Name")
        End Set
    End Property


    Private _DataTables As ObservableCollection(Of DataTableModel)
    Public Property DataTables As ObservableCollection(Of DataTableModel)
        Get
            Return Me._DataTables
        End Get
        Set(value As ObservableCollection(Of DataTableModel))
            Me._DataTables = value
            RaisePropertyChanged("DataTables")
        End Set
    End Property


    Private _IsChecked As Boolean
    Public Property IsChecked As Boolean Implements TreeViewInterface.IsChecked
        Get
            Return Me._IsChecked
        End Get
        Set(value As Boolean)
            Me._IsChecked = value
            RaisePropertyChanged("IsChecked")
            Call CheckingChildren(Of DataTableModel)(Me.DataTables, value)
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


    Public Sub MemberCheck()
        '
        If String.IsNullOrEmpty(Me.Name) Then
            Me.Name = vbNullString
        End If

        '
        If Me.DataTables Is Nothing Then
            Me.DataTables = New ObservableCollection(Of DataTableModel)
        End If

        '
        If Me.IsChecked = Nothing Then
            Me.IsChecked = False
        End If

        '
        If Me.IsEnabled = Nothing Then
            Me.IsEnabled = False
        End If
    End Sub

    Sub New()
        Call Me.MemberCheck()
    End Sub
End Class
