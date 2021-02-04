Public Class BasUpdateCommand : Inherits RpaCommandBase
    ' 2021/02/02 : アップデート実行コマンドとして作成

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer

    End Function

    Private Function Check(ByRef dat As RpaDataWrapper) As Integer

    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
        Me.ExecutableParameterCount = {0, 999}
    End Sub
End Class
