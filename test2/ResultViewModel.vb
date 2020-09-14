Public Class ResultViewModel
    Inherits ProjectBaseViewModel(Of ResultViewModel)

    Private _ResultString As String
    Public Property ResultString As String
        Get
            Return Me._ResultString
        End Get
        Set(value As String)
            Me._ResultString = value
            RaisePropertyChanged("ResultString")
        End Set
    End Property

    Public Sub AddLine(ByVal txt As String)
        Me.ResultString &= vbCrLf & txt
    End Sub


    Sub New(ByRef m As Model,
            ByRef vm As ViewModel)

        ' ビューモデルの設定
        Initializing(m, vm)
    End Sub
End Class
