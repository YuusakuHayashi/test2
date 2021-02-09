Public Interface RpaCommandInterface
    Function Execute() As Integer
    Function CanExecute() As Boolean
    ReadOnly Property ExecutableProjectArchitectures As String()
    ReadOnly Property ExecutableParameterCount As Integer()
    ReadOnly Property ExecutableUser As String()
    ReadOnly Property ExecuteIfNoProject As Boolean
End Interface
