Public Class Window2
    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        'Me.DataContext = New MyCommand.WinVM()

    End Sub
End Class

'Private Sub Window2_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
'    Dim bd As New test2.testmodule.BData

'    Try
'        test2.testmodule.RequirementCheck()
'    Catch ex As System.Exception
'        MsgBox("この処理を実行可能なシステム要件が満たされていません" & vbCrLf & "終了してください")
'        Exit Sub
'    End Try

'    test2.testmodule.defaultWinHeight = Me.Height

'    '今まで、bdはサブルーチン内で書き換えられると思っていたが、
'    '書き換えられなかったので、関数の形式にした
'    '(書き換えられないのは、配列が参照型のせいで、
'    '参照先がスコープを抜けたことで、解放されるからかもしれない)
'    bd = test2.testmodule.Load(bd)

'    '入力された値を優先
'    If test2.Window3.ServerName IsNot Nothing Then
'        If test2.Window3.ServerName <> vbNullString Then
'            Me.ServerName.Text = test2.Window3.ServerName
'        End If
'    End If
'    If test2.Window3.DBName IsNot Nothing Then
'        If test2.Window3.DBName <> vbNullString Then
'            Me.DBName.Text = test2.Window3.DBName
'        End If
'    End If
'    If test2.Window3.FieldName IsNot Nothing Then
'        If test2.Window3.FieldName <> vbNullString Then
'            Me.FieldName.Text = test2.Window3.FieldName
'        End If
'    End If
'    If test2.Window3.SrcValue IsNot Nothing Then
'        If test2.Window3.SrcValue <> vbNullString Then
'            Me.SrcValue.Text = test2.Window3.SrcValue
'        End If
'    End If
'    If test2.Window3.DistValue IsNot Nothing Then
'        If test2.Window3.DistValue <> vbNullString Then
'            Me.DistValue.Text = test2.Window3.DistValue
'        End If
'    End If
'End Sub

'Private Sub back_Click(sender As Object, e As RoutedEventArgs) Handles Back.Click
'    Call test2.testmodule.swichPages(Me, sender)
'End Sub

'Private Sub next_Click(sender As Object, e As RoutedEventArgs) Handles [Next].Click
'    Dim extFlg As System.Boolean : extFlg = False
'    Dim ck2 As Microsoft.VisualBasic.MsgBoxResult
'    Dim bd As New test2.testmodule.BData

'    'コントロール制御
'    Call Me.ControlsDisabled()

'    bd.AccessCk = Constants.vbTrue

'    If String.IsNullOrEmpty(Me.ServerName.Text) Then
'        Me.ServerErr.Content = "サーバ名が入力されていません"
'        bd.AccessCk = Constants.vbFalse
'    End If
'    bd.ServerName = Me.ServerName.Text

'    If String.IsNullOrEmpty(Me.DBName.Text) Then
'        Me.DBErr.Content = "DB名が入力されていません"
'        bd.AccessCk = Constants.vbFalse
'    End If
'    bd.DBName = Me.DBName.Text

'    If Not bd.AccessCk Then
'        Call Me.ControlsEnabled()
'        Exit Sub
'    End If

'    If String.IsNullOrEmpty(Me.FieldName.Text) Then
'        Me.FieldErr.Content = "複製対象のフィールド名が入力されていません"
'        bd.AccessCk = Constants.vbFalse
'    End If
'    bd.SrcValue = Me.SrcValue.Text

'    If String.IsNullOrEmpty(bd.SrcValue) Then
'        Me.ValueErr.Content = "複製元の値が入力されていません"
'        bd.AccessCk = Constants.vbFalse
'    End If
'    bd.DistValue = Me.DistValue.Text

'    If String.IsNullOrEmpty(bd.DistValue) Then
'        Me.ValueErr.Content = "複製先の値が入力されていません"
'        bd.AccessCk = Constants.vbFalse
'    End If
'    bd.FieldName = Me.FieldName.Text


'    bd = testmodule.AccessTest(bd)
'    If Not bd.AccessCk Then
'        Me.ServerErr.Content = "接続に失敗しました。"
'    End If

'    If Not bd.AccessCk Then
'        Call Me.ControlsEnabled()
'        Exit Sub
'    End If

'    bd = testmodule.myTableList(bd)
'    If Not bd.AccessCk Then
'        Me.DBErr.Content = "テーブル一覧の取得に失敗しました"
'    End If

'    If Not bd.AccessCk Then
'        Call Me.ControlsEnabled()
'        Exit Sub
'    End If

'    ck2 = MsgBox("実行します。よろしいですか？",
'                 Microsoft.VisualBasic.MsgBoxStyle.OkCancel)


'    'ここどうにかならないのか・・・(要変更)
'    If ck2 = MsgBoxResult.Ok Then
'        test2.Window3.ServerName = bd.ServerName
'        test2.Window3.DBName = bd.DBName
'        test2.Window3.FieldName = bd.FieldName
'        test2.Window3.SrcValue = bd.SrcValue
'        test2.Window3.DistValue = bd.DistValue
'        test2.Window3.TestDT = bd.TestDT
'        test2.Window3.SqlVer = bd.SqlVer
'        test2.Window3.TableList = bd.TableList
'        Call test2.testmodule.swichPages(Me, sender)
'    End If

'    'コントロール制御戻し
'    Call Me.ControlsEnabled()

'End Sub

'Private Sub ServerName_GotFocus(sender As Object, e As RoutedEventArgs) Handles ServerName.GotFocus
'    Try
'        Me.ServerName.SelectAll()
'    Catch ex As Exception
'    End Try
'End Sub

'Private Sub DBName_GotFocus(sender As Object, e As RoutedEventArgs) Handles DBName.GotFocus
'    Try
'        Me.DBName.SelectAll()
'    Catch ex As Exception

'    End Try
'End Sub

'Private Sub FieldName_GotFocus(sender As Object, e As RoutedEventArgs) Handles FieldName.GotFocus
'    Try
'        Me.FieldName.SelectAll()
'    Catch ex As Exception

'    End Try
'End Sub

'Private Sub SrcValue_GotFocus(sender As Object, e As RoutedEventArgs) Handles SrcValue.GotFocus
'    Try
'        Me.SrcValue.SelectAll()
'    Catch ex As Exception
'    End Try
'End Sub

'Private Sub DistValue_GotFocus(sender As Object, e As RoutedEventArgs) Handles DistValue.GotFocus
'    Try
'        Me.DistValue.SelectAll()
'    Catch ex As Exception
'    End Try
'End Sub

'Private Sub Window2_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
'End Sub

'Private Sub Window2_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
'    Me.Theme.FontSize = test2.testmodule.changeFontSize(36, e.NewSize.Height)
'    Me.Contents.FontSize = test2.testmodule.changeFontSize(20, e.NewSize.Height)
'    Me.ServerLabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.ServerName.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.ServerErr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.DBLabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.DBName.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.DBErr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.FieldLabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.FieldName.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.FieldErr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.ValueLabel.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.DistValue.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.SrcValue.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.ValueErr.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'    Me.FromTo.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
'End Sub

'Private Sub ControlsEnabled()
'    Try
'        Me.ServerName.IsEnabled = Constants.vbTrue
'        Me.DBName.IsEnabled = Constants.vbTrue
'        Me.FieldName.IsEnabled = Constants.vbTrue
'        Me.SrcValue.IsEnabled = Constants.vbTrue
'        Me.DistValue.IsEnabled = Constants.vbTrue
'        Me.Next.IsEnabled = Constants.vbTrue
'        Me.Back.IsEnabled = Constants.vbTrue
'    Catch ex As Exception
'    End Try
'End Sub

'Private Sub ControlsDisabled()
'    Try
'        Me.ServerName.IsEnabled = Constants.vbFalse
'        Me.DBName.IsEnabled = Constants.vbFalse
'        Me.FieldName.IsEnabled = Constants.vbFalse
'        Me.SrcValue.IsEnabled = Constants.vbFalse
'        Me.DistValue.IsEnabled = Constants.vbFalse
'        Me.Next.IsEnabled = Constants.vbFalse
'        Me.Back.IsEnabled = Constants.vbFalse
'    Catch ex As Exception
'    End Try
'End Sub
