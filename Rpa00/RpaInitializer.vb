Public Class RpaInitializer : Inherits JsonHandler(Of RpaInitializer)
    Private _CurrentSolutionArchitecture As String
    Public Property CurrentSolutionArchitecture As String
        Get
            Return Me._CurrentSolutionArchitecture
        End Get
        Set(value As String)
            Me._CurrentSolutionArchitecture = value
        End Set
    End Property

    Private _CurrentSolution As String
    Public Property CurrentSolution As String
        Get
            Return Me._CurrentSolution
        End Get
        Set(value As String)
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

    Private _Solutions As List(Of String)
    Public Property Solutions As List(Of String)
        Get
            If Me._Solutions Is Nothing Then
                Me._Solutions = New List(Of String)
            End If
            Return Me._Solutions
        End Get
        Set(value As List(Of String))
            Me._Solutions = value
        End Set
    End Property
End Class
