Module TreeViewModule

    '親のチェックがＯＮ（ＯＦＦ）の時、子要素のチェックをＯＮ（ＯＦＦ）にする
    Public Sub CheckingChildren(Of T As TreeViewInterface)(ByRef col As IEnumerable(Of T), ByVal b As Boolean)
        If col IsNot Nothing Then
            For Each elm In col
                elm.IsChecked = b
                elm.IsEnabled = b
            Next
        End If
    End Sub
End Module
