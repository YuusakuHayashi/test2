Class DBExplorerView
    Sub New(ByRef dbevm As DBExplorerViewModel)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Me.DataContext = dbevm
    End Sub
End Class
