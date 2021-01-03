Public MustInherit Class RpaUtilityCommandBase : Implements RpaCommandInterface

    Public Overridable ReadOnly Property ExecutableProjectArchitectures As String() Implements RpaCommandInterface.ExecutableProjectArchitectures
        Get
            Return {"AllArchType"}
        End Get
    End Property

    Public Overridable ReadOnly Property ExecutableParameterCount As Integer() Implements RpaCommandInterface.ExecutableParameterCount
        Get
            Return {0, 0}
        End Get
    End Property

    Public Overridable ReadOnly Property ExecutableUser As String() Implements RpaCommandInterface.ExecutableUser
        Get
            Return {"AllUser"}
        End Get
    End Property

    Public Overridable ReadOnly Property ExecuteIfNoProject As Boolean Implements RpaCommandInterface.ExecuteIfNoProject
        Get
            Return False
        End Get
    End Property

    Public Overloads Function Execute() As Integer Implements RpaCommandInterface.Execute
        Throw New NotImplementedException()
    End Function

    Public MustOverride Overloads Function Execute(ByRef dat As Object) As Integer


    Public Overloads Function CanExecute() As Boolean Implements RpaCommandInterface.CanExecute
        Throw New NotImplementedException()
    End Function

    Public MustOverride Overloads Function CanExecute(ByRef dat As Object) As Boolean
End Class
