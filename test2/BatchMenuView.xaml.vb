Class BatchMenuView

    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Dim bmvm As New BatchMenuViewModel
        Me.DataContext = bmvm
    End Sub
End Class
