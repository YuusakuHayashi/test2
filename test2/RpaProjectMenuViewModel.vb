Imports test2

Public Class RpaProjectMenuViewModel : Inherits BaseViewModel2

    Private _StartupIcon As BitmapImage
    Public Property StartupIcon As BitmapImage
        Get
            Return Me._StartupIcon
        End Get
        Set(value As BitmapImage)
            Me._StartupIcon = value
            RaisePropertyChanged("StartupIcon")
        End Set
    End Property

    Private _DownloadIcon As BitmapImage
    Public Property DownloadIcon As BitmapImage
        Get
            Return Me._DownloadIcon
        End Get
        Set(value As BitmapImage)
            Me._DownloadIcon = value
            RaisePropertyChanged("DownloadIcon")
        End Set
    End Property

    Public Sub _ViewInitialize()
        Dim f As Func(Of String, BitmapImage)
        f = Function(ByVal img As String) As BitmapImage
                Dim bi = New BitmapImage
                bi.BeginInit()
                bi.UriSource = New Uri(
                    img,
                    UriKind.Absolute
                )
                bi.EndInit()
                Return bi
            End Function

        Me.StartupIcon = f(AppDirectoryModel.AppImageDirectory & "\RpaProject\startup.png")
        Me.DownloadIcon = f(AppDirectoryModel.AppImageDirectory & "\RpaProject\download.png")
    End Sub

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        InitializeHandler = AddressOf _ViewInitialize
        BaseInitialize(app, vm)
    End Sub
End Class
