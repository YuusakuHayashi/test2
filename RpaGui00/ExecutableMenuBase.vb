Imports Newtonsoft.Json
Public MustInherit Class ExecutableMenuBase(Of T As {New}) : Inherits ViewModelBase(Of T)

    ' !!! RpaDataWrapper型のdllオブジェクトがキャストされる
    ' 別プロジェクトのせいか、RpaDataWrapper型としては定義できなかった
    ' (するとキャスト時エラーになる)    
    'Private _Data As Object
    '<JsonIgnore>
    'Public Property Data As Object
    '    Get
    '        Return Me._Data
    '    End Get
    '    Set(value As Object)
    '        Me._Data = value
    '    End Set
    'End Property

    'Public Overridable Sub Initialize(ByRef dat As Object)
    '    Me.Data = dat
    'End Sub
End Class
