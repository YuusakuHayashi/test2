Class MigraterView
    Sub New(ByRef mvm As MigraterViewModel)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Me.DataContext = mvm
    End Sub
End Class
