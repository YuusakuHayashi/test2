Public Class Window2
    Dim test As test2.testclass

    Private Sub Window2_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Dim bd As New test2.testmodule.BData

        test = New testclass
        test.w = Me

        test2.testmodule.defaultWinHeight = Me.Height

        '今まで、bdはサブルーチン内で書き換えられると思っていたが、
        '書き換えられなかったので、関数の形式にした
        '(書き換えられないのは、配列が参照型のせいで、
        '参照先がスコープを抜けたことで、解放されるからかもしれない)
        bd = test2.testmodule.Load(bd)
        If bd.serverName IsNot Nothing Then
            Me.srvtxt.Text = bd.serverName
        End If
        If bd.databaseName IsNot Nothing Then
            Me.dbtxt.Text = bd.databaseName
        End If
        If bd.fieldName IsNot Nothing Then
            Me.fieldtxt.Text = bd.fieldName
        End If
        If bd.sourceValue IsNot Nothing Then
            Me.beforetxt.Text = bd.sourceValue
        End If
        If bd.distinationValue IsNot Nothing Then
            Me.aftertxt.Text = bd.distinationValue
        End If

        If test2.Window3.serverName IsNot Nothing Then
            If test2.Window3.serverName <> vbNullString Then
                Me.srvtxt.Text = test2.Window3.serverName
            End If
        End If
        If test2.Window3.databaseName IsNot Nothing Then
            If test2.Window3.databaseName <> vbNullString Then
                Me.dbtxt.Text = test2.Window3.databaseName
            End If
        End If
        If test2.Window3.fieldName IsNot Nothing Then
            If test2.Window3.fieldName <> vbNullString Then
                Me.fieldtxt.Text = test2.Window3.fieldName
            End If
        End If
        If test2.Window3.sourceValue IsNot Nothing Then
            If test2.Window3.sourceValue <> vbNullString Then
                Me.beforetxt.Text = test2.Window3.sourceValue
            End If
        End If
        If test2.Window3.distinationValue IsNot Nothing Then
            If test2.Window3.distinationValue <> vbNullString Then
                Me.aftertxt.Text = test2.Window3.distinationValue
            End If
        End If
    End Sub

    Private Sub back_Click(sender As Object, e As RoutedEventArgs) Handles back.Click
        Call test2.testmodule.swichPages(test.w, sender)
    End Sub

    Private Sub next_Click(sender As Object, e As RoutedEventArgs) Handles [next].Click
        Dim extFlg As System.Boolean : extFlg = False
        Dim ck As System.Boolean    'access test
        Dim ck2 As MsgBoxResult   'confirm
        Dim bd As New test2.testmodule.BData

        bd.serverName = srvtxt.Text
        bd.databaseName = dbtxt.Text

        If bd.serverName = Constants.vbNullString Then
            Me.srverr.Content = "サーバ名が入力されていません"
            extFlg = Constants.vbTrue
        End If

        If bd.databaseName = Constants.vbNullString Then
            Me.dberr.Content = "DB名が入力されていません"
            extFlg = Constants.vbTrue
        End If

        If extFlg Then
            Exit Sub
        End If

        ck = testmodule.testAccess(bd)
        If Not ck Then
            Me.dberr.Content = "接続に失敗しました。" & "サーバ名・データベース名に誤りがないかご確認ください"
            Exit Sub
        End If


        bd.sourceValue = beforetxt.Text
        bd.distinationValue = aftertxt.Text
        bd.fieldName = fieldtxt.Text

        If bd.fieldName = Constants.vbNullString Then
            Me.fielderr.Content = "複製対象のフィールド名が入力されていません"
            extFlg = Constants.vbTrue
        End If

        If bd.sourceValue = Constants.vbNullString Then
            Me.valueerr.Content = "複製元の値が入力されていません"
            extFlg = Constants.vbTrue
        End If

        If bd.distinationValue = Constants.vbNullString Then
            Me.valueerr.Content = "複製先の値が入力されていません"
            extFlg = Constants.vbTrue
        End If

        If extFlg Then
            Exit Sub
        End If

        ck2 = MsgBox("実行します。よろしいですか？", Microsoft.VisualBasic.MsgBoxStyle.OkCancel)

        'ここどうにかならないのか・・・
        If ck2 = MsgBoxResult.Ok Then
            test2.Window3.serverName = bd.serverName
            test2.Window3.databaseName = bd.databaseName
            test2.Window3.fieldName = bd.fieldName
            test2.Window3.sourceValue = bd.sourceValue
            test2.Window3.distinationValue = bd.distinationValue
            Call test2.testmodule.swichPages(test.w, sender)
        End If

        bd = Nothing
    End Sub

    Private Sub srvtxt_GotFocus(sender As Object, e As RoutedEventArgs) Handles srvtxt.GotFocus
        Me.srvtxt.SelectAll()
    End Sub

    Private Sub dbtxt_GotFocus(sender As Object, e As RoutedEventArgs) Handles dbtxt.GotFocus
        Me.dbtxt.SelectAll()
    End Sub

    Private Sub fieldtxt_GotFocus(sender As Object, e As RoutedEventArgs) Handles fieldtxt.GotFocus
        Me.fieldtxt.SelectAll()
    End Sub

    Private Sub beforetxt_GotFocus(sender As Object, e As RoutedEventArgs) Handles beforetxt.GotFocus
        Me.beforetxt.SelectAll()
    End Sub

    Private Sub aftertxt_GotFocus(sender As Object, e As RoutedEventArgs) Handles aftertxt.GotFocus
        Me.aftertxt.SelectAll()
    End Sub

    Private Sub Window2_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
    End Sub

    Private Sub Window2_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Me.srvlabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.srvtxt.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.srverr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.dblabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.dbtxt.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.dberr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.fieldlabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.fieldtxt.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.fielderr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.value.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.aftertxt.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.beforetxt.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.valueerr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
        Me.label2.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        Me.label3.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        Me.label4.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        Me.label5.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        Me.label6.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        Me.label7.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
    End Sub
End Class
