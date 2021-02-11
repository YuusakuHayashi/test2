Public Class ExecuteMenu
    Private _Icon As BitmapImage
    Public Property Icon As BitmapImage
        Get
            Return Me._Icon
        End Get
        Set(value As BitmapImage)
            Me._Icon = value
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

    Private _Content As Object
    Public Property Content As Object
        Get
            Return Me._Content
        End Get
        Set(value As Object)
            Me._Content = value
        End Set
    End Property

    Private _ContentType As Type
    Public Property ContentType As Type
        Get
            Return Me._ContentType
        End Get
        Set(value As Type)
            Me._ContentType = value
        End Set
    End Property
End Class
