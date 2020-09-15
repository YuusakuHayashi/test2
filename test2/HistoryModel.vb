Public Class HistoryModel : Inherits BaseModel
    Private _Contents As String
    Public Property Contents As String
        Get
            Return Me._Contents
        End Get
        Set(value As String)
            Me._Contents = value
            RaisePropertyChanged("Contents")
        End Set
    End Property

    Public Sub NewLine(ByVal txt As String)
        Me.Contents = txt
    End Sub

    Public Sub AddLine(ByVal txt As String)
        Me.Contents &= vbCrLf & txt
    End Sub

    Sub New()
        Me.NewLine("Historyは初期化されました")
    End Sub
End Class
