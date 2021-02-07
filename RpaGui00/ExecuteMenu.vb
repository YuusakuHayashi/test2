Public Class ExecuteMenu : Inherits ViewModelBase

    Private _Name As String
    Public Property Name As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
        End Set
    End Property

    Private _IsSelected As Boolean
    Public Property IsSelected As Boolean
        Get
            Return Me._IsSelected
        End Get
        Set(value As Boolean)
            Me._IsSelected = value
            RaisePropertyChanged("IsSelected")
        End Set
    End Property

    'Public Delegate Function ExecuteDelegater(ByRef dat As Object) As Integer
    'Private _ExecuteHandler As ExecuteDelegater
    'Public Property ExecuteHandler As ExecuteDelegater
    '    Get
    '        Return Me._ExecuteHandler
    '    End Get
    '    Set(value As ExecuteDelegater)
    '        Me._ExecuteHandler = value
    '    End Set
    'End Property

    Private _DataObject As Object

    Public Property DataObject As Object
        Get
            Return Me._DataObject
        End Get
        Set(value As Object)
            Me._DataObject = value
        End Set
    End Property
End Class
