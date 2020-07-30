Public Class MainWin
    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Dim cfvm As New ConfigFileVM
        Me.ConfigFileValue.DataContext = cfvm


        Dim pm As New ProjectModel
        Me.TreeView.DataContext = New TreeViewVM(pm.LoadTreeViewModel)


        Me.AccessButton.DataContext = New AccessButtonVM(cfvm)
    End Sub
End Class
