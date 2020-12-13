Public Class RpaPackage
    Private _Name As String
    Public Property Name As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
        End Set
    End Property

    Private _Latest As Boolean
    Public Property Latest As Boolean
        Get
            Return Me._Latest
        End Get
        Set(value As Boolean)
            Me._Latest = value
        End Set
    End Property

    Private _UpdateInfos As List(Of RpaPackage.UpdateInfo)
    Public Property UpdateInfos As List(Of RpaPackage.UpdateInfo)
        Get
            Return Me._UpdateInfos
        End Get
        Set(value As List(Of RpaPackage.UpdateInfo))
            Me._UpdateInfos = value
        End Set
    End Property

    Public Class UpdateInfo
        Private _SourceFile As String
        Public Property SourceFile As String
            Get
                Return Me._SourceFile
            End Get
            Set(value As String)
                Me._SourceFile = value
            End Set
        End Property

        Private _DistinationSubDirectory As String
        Public Property DistinationSubDirectory As String
            Get
                Return Me._DistinationSubDirectory
            End Get
            Set(value As String)
                Me._DistinationSubDirectory = value
            End Set
        End Property

        Private _Author As String
        Public Property Author As String
            Get
                Return Me._Author
            End Get
            Set(value As String)
                Me._Author = value
            End Set
        End Property
    End Class
End Class
