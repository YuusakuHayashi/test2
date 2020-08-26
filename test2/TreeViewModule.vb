Module TreeViewModule

    '親のチェックがＯＮ（ＯＦＦ）の時、子要素のチェックをＯＮ（ＯＦＦ）にする
    Sub CheckingChildren(Of T As TreeViewInterface)(ByRef col As IEnumerable(Of T), ByVal b As Boolean)
        For Each elm In col
            elm.IsChecked = b
            elm.IsEnabled = b
        Next
    End Sub

End Module
