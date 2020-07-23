Imports System.Collections.ObjectModel

'Public Class Model
'    Public Name As String
'    Public Server As ServerModel
'End Class

Public Class ServerModel
    Public ServerName As String
End Class

Public Class DataBase
    Public DataBaseName As String
    Public DataBaseChild As List(Of DataBase)
End Class

Public Class DataTable
    Public DataTableName As String
    Public DataTableChild As List(Of DataTable)
End Class

Public Class Person
    Public Name As String
    Public Child As ObservableCollection(Of Person)
End Class

Public Class Model
    'Private _Name As String
    'Private _Functions As List(Of Plugin)

    Private _ServerName As String
    Public Property ServerName As String
        Get
            Return _ServerName
        End Get
        Set(value As String)
            _ServerName = value
        End Set
    End Property

    Private Property _DataBaseName As String
    Public Property DataBaseName As String
        Get
            Return _DataBaseName
        End Get
        Set(value As String)
            _DataBaseName = value
        End Set
    End Property

    Private Property _DataTableName As String
    Public Property DataTableName As String
        Get
            Return _DataTableName
        End Get
        Set(value As String)
            _DataTableName = value
        End Set
    End Property

    Private _Child As List(Of Model)
    Public Property Child As List(Of Model)
        Get
            Return _Child
        End Get
        Set
            _Child = Value
        End Set
    End Property
End Class