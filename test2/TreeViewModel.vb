Public Class TreeViewModel
    'TreeView は考え方が複雑・・・
    '描画線用のデータ構造ではTreeViewのリスト型
    Private Property _RealName As String
    Public Property RealName As String
        Get
            Return _RealName
        End Get
        Set(value As String)
            _RealName = value
        End Set
    End Property

    Private _Child As List(Of TreeViewModel)
    Public Property Child As List(Of TreeViewModel)
        Get
            Return _Child
        End Get
        Set(value As List(Of TreeViewModel))
            _Child = value
        End Set
    End Property
    '-----------------------------------------------------------------------------'
End Class
