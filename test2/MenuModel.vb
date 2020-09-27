Public Class MenuModel : Inherits BaseModel
    Private _ViewName As String
    Public Property ViewName As String
        Get
            Return Me._ViewName
        End Get
        Set(value As String)
            Me._ViewName = value
        End Set
    End Property


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


    Private _DisplayName As String
    Public Property DisplayName As String
        Get
            Return Me._DisplayName
        End Get
        Set(value As String)
            Me._DisplayName = value
            RaisePropertyChanged("DisplayName")
        End Set
    End Property


    Private _IsSelected As Boolean
    Public Property IsSelected As Boolean
        Get
            Return Me._IsSelected
        End Get
        Set(value As Boolean)
            Me._IsSelected = value
            If value Then
                MenuChangedEventListener.Instance.RaiseChangeRequested(Me)
            End If
        End Set
    End Property

    'Private _IsExpanded As Boolean
    'Public Property IsExpanded As Boolean
    '    Get
    '        Return Me._IsExpanded
    '    End Get
    '    Set(value As Boolean)
    '        Me._IsExpanded = value
    '        RaisePropertyChanged("IsExpanded")
    '    End Set
    'End Property
End Class
