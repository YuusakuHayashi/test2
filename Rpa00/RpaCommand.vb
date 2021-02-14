Public Class RpaCommand
    Private _ProjectName As String
    Public Property ProjectName As String
        Get
            Return Me._ProjectName
        End Get
        Set(value As String)
            Me._ProjectName = value
        End Set
    End Property

    Private _ProjectArchTypeName As String
    Public Property ProjectArchTypeName As String
        Get
            Return Me._ProjectArchTypeName
        End Get
        Set(value As String)
            Me._ProjectArchTypeName = value
        End Set
    End Property

    Private _RobotName As String
    Public Property RobotName As String
        Get
            Return Me._RobotName
        End Get
        Set(value As String)
            Me._RobotName = value
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

    Private _Result As String
    Public Property Result As String
        Get
            Return Me._Result
        End Get
        Set(value As String)
            Me._Result = value
        End Set
    End Property

    Private _Message As String
    Public Property Message As String
        Get
            Return Me._Message
        End Get
        Set(value As String)
            Me._Message = value
        End Set
    End Property

    Private _UserName As String
    Public Property UserName As String
        Get
            Return Me._UserName
        End Get
        Set(value As String)
            Me._UserName = value
        End Set
    End Property

    Private _RunDate As Date
    Public Property RunDate As Date
        Get
            Return Me._RunDate
        End Get
        Set(value As Date)
            Me._RunDate = value
        End Set
    End Property

    Private _ExecuteTime As TimeSpan
    Public Property ExecuteTime As TimeSpan
        Get
            Return Me._ExecuteTime
        End Get
        Set(value As TimeSpan)
            Me._ExecuteTime = value
        End Set
    End Property

    Private _IsAutoCommand As Boolean
    Public Property IsAutoCommand As Boolean
        Get
            Return Me._IsAutoCommand
        End Get
        Set(value As Boolean)
            Me._IsAutoCommand = value
        End Set
    End Property

    Private _UtilityCommand As String
    Public Property UtilityCommand As String
        Get
            Return Me._UtilityCommand
        End Get
        Set(value As String)
            Me._UtilityCommand = value
        End Set
    End Property

    Private _TrueCommand As String
    Public Property TrueCommand As String
        Get
            Return Me._TrueCommand
        End Get
        Set(value As String)
            Me._TrueCommand = value
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

    Private _ParametersText As String
    Public Property ParametersText As String
        Get
            Return Me._ParametersText
        End Get
        Set(value As String)
            Me._ParametersText = value
        End Set
    End Property
End Class
