Public Class MyProjectModel : Inherits ProjectInfoModel
    Private _ProjectName As String
    Protected Property ProjectName As String
        Get
            Return _ProjectName
        End Get
        Set(value As String)
            _ProjectName = value
        End Set
    End Property
End Class
