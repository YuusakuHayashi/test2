Class MainWindow
    Dim test As testclass


    Private Sub MainWindow_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        test = New testclass
        test.w = Me

        test2.testmodule.defaultWinHeight = Me.Height
    End Sub

    Private Sub next_Click(sender As Object, e As RoutedEventArgs) Handles [next].Click
        Call testmodule.swichPages(test.w, sender)
    End Sub

    Private Sub MainWindow_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Me.trn1.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
    End Sub
End Class
