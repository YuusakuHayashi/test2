Imports System.Collections.ObjectModel

Public Class DBTestModel 
    Inherits BaseModel(Of DBTestModel)
    Public ServerName As String
    Public DataBaseName As String
    Public ConnectionString As String
    Public Server As ServerModel
    Public History As HistoryModel
    'Sub New()
    '    Server = New ServerModel With {
    '        .DataBases = New ObservableCollection(Of DataBaseModel) From {
    '            New DataBaseModel With {.Name = "hoge"},
    '            New DataBaseModel With {.Name = "fuga"},
    '            New DataBaseModel With {.Name = "piyo"}
    '        }
    '    }
    'End Sub
End Class
