Public Class RpaDataWrapper
    Private _Project As Object
    Public Property Project As Object
        Get
            Return Me._Project
        End Get
        Set(value As Object)
            Me._Project = value
        End Set
    End Property

    Private _Transaction As RpaTransaction
    Public Property Transaction As RpaTransaction
        Get
            If Me._Transaction Is Nothing Then
                Me._Transaction = New RpaTransaction
            End If
            Return Me._Transaction
        End Get
        Set(value As RpaTransaction)
            Me._Transaction = value
        End Set
    End Property

    Private _Initializer As RpaInitializer
    Public Property Initializer As RpaInitializer
        Get
            If Me._Initializer Is Nothing Then
                Me._Initializer = New RpaInitializer
            End If
            Return Me._Initializer
        End Get
        Set(value As RpaInitializer)
            Me._Initializer = value
        End Set
    End Property

    Private _System As RpaSystem
    Public Property System As RpaSystem
        Get
            If Me._System Is Nothing Then
                Me._System = New RpaSystem
            End If
            Return Me._System
        End Get
        Set(value As RpaSystem)
            Me._System = value
        End Set
    End Property
End Class
