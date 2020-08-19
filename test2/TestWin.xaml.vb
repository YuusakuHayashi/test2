Public Class TestWin
    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Dim tvm As New TreeViewModel
        Me.treeView.DataContext = tvm
    End Sub
End Class
