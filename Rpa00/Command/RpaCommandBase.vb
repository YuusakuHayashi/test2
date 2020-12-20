Public MustInherit Class RpaCommandBase
    Public MustOverride Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object) As Integer
    Public MustOverride ReadOnly Property ExecutableProjectArchitectures As Integer()
    Public MustOverride ReadOnly Property CanExecute(trn As RpaTransaction, rpa As Object) As Boolean
    Public MustOverride ReadOnly Property ExecutableParameterCount() As Integer()
    Public MustOverride ReadOnly Property ExecutableUserLevel() As Integer
End Class
