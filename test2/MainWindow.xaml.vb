Class MainWindow
    Private Sub MainWindow_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        test2.testmodule.defaultWinHeight = Me.Height
    End Sub

    Private Sub Next_Click(sender As Object, e As RoutedEventArgs) Handles [Next].Click
        Call testmodule.swichPages(Me, sender)
    End Sub

    Private Sub MainWindow_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Me.Title.FontSize = test2.testmodule.changeFontSize(36, e.NewSize.Height)
        Me.trn1.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
    End Sub
End Class
