Public Class MainWin
    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Me.ConfigFileValue.DataContext = New ConfigFileVM
        Me.TreeView.DataContext = New TreeViewVM
        Me.AccessButton.DataContext = New AccessButtonVM
    End Sub
End Class
