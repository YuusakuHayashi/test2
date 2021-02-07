Public Class ExecutableMenuBase : Inherits ViewModelBase

    ' !!! RpaDataWrapper型のdllオブジェクトがキャストされる
    ' 別プロジェクトのせいか、RpaDataWrapper型としては定義できなかった
    ' (するとキャスト時エラーになる)
    Private _Data As Object
    Public Property Data As Object
        Get
            Return Me._Data
        End Get
        Set(value As Object)
            Me._Data = value
        End Set
    End Property
End Class
