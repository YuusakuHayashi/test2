Public Class Window2

    Private Const INITIAL_DT = "TESTDB"
    Private Const INITIAL_SQLVER = 2015

    Private Sub ControlDisable()
        Me.ServerName.IsEnabled = Constants.vbFalse
        Me.DBName.IsEnabled = Constants.vbFalse
        Me.FieldName.IsEnabled = Constants.vbFalse
        Me.SrcValue.IsEnabled = Constants.vbFalse
        Me.DistValue.IsEnabled = Constants.vbFalse
        Me.Next.IsEnabled = Constants.vbFalse
        Me.Back.IsEnabled = Constants.vbFalse
    End Sub

    Private Sub ControlEnable()
        Me.ServerName.IsEnabled = Constants.vbTrue
        Me.DBName.IsEnabled = Constants.vbTrue
        Me.FieldName.IsEnabled = Constants.vbTrue
        Me.SrcValue.IsEnabled = Constants.vbTrue
        Me.DistValue.IsEnabled = Constants.vbTrue
        Me.Next.IsEnabled = Constants.vbTrue
        Me.Back.IsEnabled = Constants.vbTrue
    End Sub



    Private Sub Window2_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Dim bd As New test2.testmodule.BData

        test2.testmodule.defaultWinHeight = Me.Height

        '今まで、bdはサブルーチン内で書き換えられると思っていたが、
        '書き換えられなかったので、関数の形式にした
        '(書き換えられないのは、配列が参照型のせいで、
        '参照先がスコープを抜けたことで、解放されるからかもしれない)
        bd = test2.testmodule.Load(bd)

        If bd.ServerName IsNot Nothing Then
            Me.ServerName.Text = bd.ServerName
        End If
        If bd.DBName IsNot Nothing Then
            Me.DBName.Text = bd.DBName
        End If
        If bd.FieldName IsNot Nothing Then
            Me.FieldName.Text = bd.FieldName
        End If
        If bd.SrcValue IsNot Nothing Then
            Me.SrcValue.Text = bd.SrcValue
        End If
        If bd.DistValue IsNot Nothing Then
            Me.DistValue.Text = bd.DistValue
        End If

        If test2.Window3.ServerName IsNot Nothing Then
            If test2.Window3.ServerName <> vbNullString Then
                Me.ServerName.Text = test2.Window3.ServerName
            End If
        End If
        If test2.Window3.DBName IsNot Nothing Then
            If test2.Window3.DBName <> vbNullString Then
                Me.DBName.Text = test2.Window3.DBName
            End If
        End If
        If test2.Window3.FieldName IsNot Nothing Then
            If test2.Window3.FieldName <> vbNullString Then
                Me.FieldName.Text = test2.Window3.FieldName
            End If
        End If
        If test2.Window3.SrcValue IsNot Nothing Then
            If test2.Window3.SrcValue <> vbNullString Then
                Me.SrcValue.Text = test2.Window3.SrcValue
            End If
        End If
        If test2.Window3.DistValue IsNot Nothing Then
            If test2.Window3.DistValue <> vbNullString Then
                Me.DistValue.Text = test2.Window3.DistValue
            End If
        End If
    End Sub

    Private Sub back_Click(sender As Object, e As RoutedEventArgs) Handles Back.Click
        Call test2.testmodule.swichPages(Me, sender)
    End Sub

    Private Sub next_Click(sender As Object, e As RoutedEventArgs) Handles [Next].Click
        Dim extFlg As System.Boolean : extFlg = False
        Dim ck As System.Boolean
        Dim ck2 As Microsoft.VisualBasic.MsgBoxResult
        Dim bd As New test2.testmodule.BData

        Call Me.ControlDisable()

        bd.ServerName = Me.ServerName.Text
        bd.DBName = Me.DBName.Text

        If bd.ServerName = Constants.vbNullString Then
            Me.ServerErr.Content = "サーバ名が入力されていません"
            extFlg = Constants.vbTrue
        End If

        If bd.DBName = Constants.vbNullString Then
            Me.DBErr.Content = "DB名が入力されていません"
            extFlg = Constants.vbTrue
        End If

        If extFlg Then
            Call Me.ControlEnable()
            Exit Sub
        End If

        ck = testmodule.AccessTest(bd)
        If Not ck Then
            Me.DBErr.Content = "接続に失敗しました。" & "サーバ名・データベース名に誤りがないかご確認ください"
            Call Me.ControlEnable()
            Exit Sub
        End If


        bd.SrcValue = Me.SrcValue.Text
        bd.DistValue = Me.DistValue.Text
        bd.FieldName = Me.FieldName.Text

        If bd.FieldName = Constants.vbNullString Then
            Me.FieldErr.Content = "複製対象のフィールド名が入力されていません"
            extFlg = Constants.vbTrue
        End If

        If bd.SrcValue = Constants.vbNullString Then
            Me.ValueErr.Content = "複製元の値が入力されていません"
            extFlg = Constants.vbTrue
        End If

        If bd.DistValue = Constants.vbNullString Then
            Me.ValueErr.Content = "複製先の値が入力されていません"
            extFlg = Constants.vbTrue
        End If

        If extFlg Then
            Call Me.ControlEnable()
            Exit Sub
        End If

        If bd.TestDT Is Nothing Then
            bd.TestDT = INITIAL_DT
        End If
        If bd.TestDT = Constants.vbNullString Then
            bd.TestDT = INITIAL_DT
        End If

        If bd.SqlVer = Nothing Then
            bd.SqlVer = INITIAL_SQLVER
        End If
        If bd.SqlVer = 0 Then
            bd.SqlVer = INITIAL_SQLVER
        End If

        ck2 = MsgBox("実行します。よろしいですか？",
                     Microsoft.VisualBasic.MsgBoxStyle.OkCancel)

        'ここどうにかならないのか・・・
        If ck2 = MsgBoxResult.Ok Then
            test2.Window3.ServerName = bd.ServerName
            test2.Window3.DBName = bd.DBName
            test2.Window3.FieldName = bd.FieldName
            test2.Window3.SrcValue = bd.SrcValue
            test2.Window3.DistValue = bd.DistValue
            test2.Window3.TestDT = bd.TestDT
            test2.Window3.SqlVer = bd.SqlVer
            Call test2.testmodule.swichPages(Me, sender)
        End If

        Call Me.ControlEnable()

        bd = Nothing
    End Sub

    Private Sub ServerName_GotFocus(sender As Object, e As RoutedEventArgs) Handles ServerName.GotFocus
        Me.ServerName.SelectAll()
    End Sub

    Private Sub DBName_GotFocus(sender As Object, e As RoutedEventArgs) Handles DBName.GotFocus
        Me.DBName.SelectAll()
    End Sub

    Private Sub FieldName_GotFocus(sender As Object, e As RoutedEventArgs) Handles FieldName.GotFocus
        Me.FieldName.SelectAll()
    End Sub

    Private Sub SrcValue_GotFocus(sender As Object, e As RoutedEventArgs) Handles SrcValue.GotFocus
        Me.SrcValue.SelectAll()
    End Sub

    Private Sub DistValue_GotFocus(sender As Object, e As RoutedEventArgs) Handles DistValue.GotFocus
        Me.DistValue.SelectAll()
    End Sub

    Private Sub Window2_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
    End Sub

    Private Sub Window2_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Me.Progress.FontSize = test2.testmodule.changeFontSize(36, e.NewSize.Height)
        Me.Contents.FontSize = test2.testmodule.changeFontSize(36, e.NewSize.Height)
        Me.ServerLabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.ServerName.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.ServerErr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.DBLabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.DBName.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.DBErr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.FieldLabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.FieldName.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.FieldErr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.ValueLabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.DistValue.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.SrcValue.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.ValueErr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.FromTo.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        'Me.label2.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        'Me.label3.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        'Me.label4.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        'Me.label5.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        'Me.label6.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        'Me.label7.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
    End Sub

End Class
