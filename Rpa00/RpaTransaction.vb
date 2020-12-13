Public Class RpaTransaction
    Private _ReturnCode As Integer
    Public Property ReturnCode As Integer
        Get
            Return Me._ReturnCode
        End Get
        Set(value As Integer)
            Me._ReturnCode = value
        End Set
    End Property

    Private _CommandText As String
    Public Property CommandText As String
        Get
            Return Me._CommandText
        End Get
        Set(value As String)
            Me._CommandText = value
        End Set
    End Property

    Private _MainCommand As String
    Public Property MainCommand As String
        Get
            Return Me._MainCommand
        End Get
        Set(value As String)
            Me._MainCommand = value
        End Set
    End Property

    Private _Parameters As List(Of String)
    Public Property Parameters As List(Of String)
        Get
            If Me._Parameters Is Nothing Then
                Me._Parameters = New List(Of String)
            End If
            Return Me._Parameters
        End Get
        Set(value As List(Of String))
            Me._Parameters = value
        End Set
    End Property

    Private _ExitFlag As Boolean
    Public Property ExitFlag As Boolean
        Get
            Return Me._ExitFlag
        End Get
        Set(value As Boolean)
            Me._ExitFlag = value
        End Set
    End Property

    Public Sub CreateCommand()
        Dim texts() As String
        Me.MainCommand = vbNullString
        Me.Parameters = New List(Of String)

        texts = Me.CommandText.Split(" ")
        Me.MainCommand = texts(0)

        For Each p In texts
            If Not p = texts.First Then
                Me.Parameters.Add(p)
            End If
        Next
    End Sub
End Class
