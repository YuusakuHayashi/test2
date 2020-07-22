Imports System.Collections.ObjectModel

Public Class Model
    Public Name As String
    Public Server As ServerModel
End Class

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
