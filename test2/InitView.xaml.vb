Class InitView

    Private _Navi As NavigationService

    Sub New(ByRef ivm As InitViewModel)
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        'Dim ivm As New InitViewModel()
        Me.DataContext = ivm
    End Sub

End Class
