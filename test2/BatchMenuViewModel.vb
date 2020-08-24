Public Class BatchMenuViewModel

    Private _Children As List(Of SqlModel)
    Public Property Children As List(Of SqlModel)
        Get
            Return Me._Children
        End Get
        Set(value As List(Of SqlModel))
            Me._Children = value
        End Set
    End Property


    Sub New()
        'TEST
        Me.Children = New List(Of SqlModel) From {
            New SqlModel With {
                .FieldName = "hoge"
            },
            New SqlModel With {
                .FieldName = "fuga"
            }
        }
    End Sub
End Class
