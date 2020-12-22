Public Class RpaInitializer : Inherits JsonHandler(Of RpaInitializer)
    Private _CurrentSolution As RpaSolution
    Public Property CurrentSolution As RpaSolution
        Get
            Return Me._CurrentSolution
        End Get
        Set(value As RpaSolution)
            Me._CurrentSolution = value
        End Set
    End Property

    Private _AutoLoad As Boolean
    Public Property AutoLoad As Boolean
        Get
            Return Me._AutoLoad
        End Get
        Set(value As Boolean)
            Me._AutoLoad = value
        End Set
    End Property

    Private _Solutions As List(Of RpaSolution)
    Public Property Solutions As List(Of RpaSolution)
        Get
            If Me._Solutions Is Nothing Then
                Me._Solutions = New List(Of RpaSolution)
            End If
            Return Me._Solutions
        End Get
        Set(value As List(Of RpaSolution))
            Me._Solutions = value
        End Set
    End Property

    Public Class RpaSolution
        Private _Architecture As Integer
        Public Property Architecture As Integer
            Get
                Return Me._Architecture
            End Get
            Set(value As Integer)
                Me._Architecture = value
            End Set
        End Property

        Private _Name As String
        Public Property Name As String
            Get
                Return Me._Name
            End Get
            Set(value As String)
                Me._Name = value
            End Set
        End Property

        Private _SolutionDirectory As String
        Public Property SolutionDirectory As String
            Get
                Return Me._SolutionDirectory
            End Get
            Set(value As String)
                Me._SolutionDirectory = value
            End Set
        End Property
    End Class
End Class
