Imports test2
Imports System.Collections.ObjectModel

Public Class RpaProjectModel
    Public RootDirectoryName As String
    Public UserDirectoryName As String
    Public SystemDirectoryName As String
    Public PythonPathName As String
    Public Rpas As ObservableCollection(Of RpaModel)
End Class
