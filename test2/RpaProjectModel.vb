Imports test2
Imports System.Collections.ObjectModel

Public Class RpaProjectModel
    Public RootDirectoryName As String
    Public UserDirectoryName As String
    Public SystemDirectoryName As String
    Public PythonPathName As String

    Private _Rpas As ObservableCollection(Of RpaModel)
    Public Property Rpas As ObservableCollection(Of RpaModel)
        Get
            If Me._Rpas Is Nothing Then
                Me._Rpas = New ObservableCollection(Of RpaModel)
            End If
            Return Me._Rpas
        End Get
        Set(value As ObservableCollection(Of RpaModel))
            Me._Rpas = value
        End Set
    End Property
    Public ReadOnly Property RootSysDirectoryName As String
        Get
            Return Me.RootDirectoryName & "\sys"
        End Get
    End Property
    Public ReadOnly Property SysDirectoryName As String
        Get
            Return Me.SystemDirectoryName & "\sys"
        End Get
    End Property
    Public ReadOnly Property MyDirFileName As String
        Get
            Return Me.SystemDirectoryName & "\mydir"
        End Get
    End Property
    Public ReadOnly Property MyPythonFileName As String
        Get
            Return Me.SystemDirectoryName & "\mypython"
        End Get
    End Property
    Public ReadOnly Property RpaProjectFileName As String
        Get
            Return Me.SystemDirectoryName & "\rpa_project.yaml"
        End Get
    End Property
End Class
