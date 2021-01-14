Public MustInherit Class RpaCommandBase : Implements RpaCommandInterface

    Public Overridable ReadOnly Property ExecutableProjectArchitectures As String() Implements RpaCommandInterface.ExecutableProjectArchitectures
        Get
            Return {"AllArchitectures"}
        End Get
    End Property

    Public Overridable ReadOnly Property ExecutableParameterCount() As Integer() Implements RpaCommandInterface.ExecutableParameterCount
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

    Public Overloads Function CanExecute() As Boolean Implements RpaCommandInterface.CanExecute
        Throw New NotImplementedException
    End Function

    Public Overridable Overloads Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        Return True
    End Function

    Public Overloads Function Execute() As Integer Implements RpaCommandInterface.Execute
        Throw New NotImplementedException()
    End Function

    Public Overridable Overloads Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Dim i As Integer = -1
        Try
            Dim exdlg As ExecuteDelegater = Me.ExecuteHandler
            i = exdlg(dat)
        Catch ex As Exception
            Console.WriteLine($"({Me.GetType.Name}) {ex.Message}")
            Console.WriteLine()
            i = -1
        End Try
        Return i
    End Function

    Public Delegate Function ExecuteDelegater(ByRef dat As RpaDataWrapper) As Integer

    Private _ExecuteHandler As ExecuteDelegater
    Public Property ExecuteHandler As ExecuteDelegater
        Get
            Return Me._ExecuteHandler
        End Get
        Set(value As ExecuteDelegater)
            Me._ExecuteHandler = value
        End Set
    End Property
End Class
