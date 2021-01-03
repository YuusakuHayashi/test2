﻿Public Class RpaTransaction
    Private _ReturnCode As Integer
    Public Property ReturnCode As Integer
        Get
            Return Me._ReturnCode
        End Get
        Set(value As Integer)
            Me._ReturnCode = value
        End Set
    End Property

    Private _ProjectArchitecture As Integer
    Public Property ProjectArchitecture As Integer
        Get
            Return Me._ProjectArchitecture
        End Get
        Set(value As Integer)
            Me._ProjectArchitecture = value
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

    Private _Modes As List(Of String)
    Public Property Modes As List(Of String)
        Get
            If Me._Modes Is Nothing Then
                Me._Modes = New List(Of String)
            End If
            Return Me._Modes
        End Get
        Set(value As List(Of String))
            Me._Modes = value
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

        Dim i As Integer = 0
        For Each p In texts
            If i <> 0 Then
                Me.Parameters.Add(p)
            End If
            i += 1
        Next
    End Sub

    Public Function ShowRpaIndicator(ByRef dat As RpaDataWrapper) As String
        If Me.Modes.Count = 0 Then
            If dat.Project IsNot Nothing Then
                Console.Write($"{dat.Project.ProjectName}\{dat.Project.RobotAlias}>")
            Else
                Console.Write("NoRpa>")
            End If
        Else
            For Each mode In Me.Modes
                If Me.Modes.Last = mode Then
                    Console.Write($"{mode}>")
                Else
                    Console.Write($"{mode}\")
                End If
            Next
        End If
        Return Console.ReadLine()
    End Function
End Class
