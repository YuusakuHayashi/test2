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

    Private _ExecutableUser As String()
    Public Property ExecutableUser As String() Implements RpaCommandInterface.ExecutableUser
        Get
            If Me._ExecutableUser Is Nothing Then
                Me._ExecutableUser = {"AllUser"}
            End If
            Return Me._ExecutableUser
        End Get
        Set(value As String())
            Me._ExecutableUser = value 
        End Set
    End Property

    Public Overridable ReadOnly Property ExecuteIfNoProject As Boolean Implements RpaCommandInterface.ExecuteIfNoProject
        Get
            Return False
        End Get
    End Property

    Public Overloads Function CanExecute() As Boolean Implements RpaCommandInterface.CanExecute
        Throw New NotImplementedException
    End Function

    'Public Overridable Overloads Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
    '    Return True
    'End Function

    Public Overridable Overloads Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        Dim i As Boolean = False
        Try
            Dim dlg As CanExecuteDelegater = Me.CanExecuteHandler
            If dlg Is Nothing Then
                i = True
            Else
                i = dlg(dat)
            End If
        Catch ex As Exception
            Console.WriteLine($"({Me.GetType.Name}.CanExecuteHandler) {ex.Message}")
            Console.WriteLine()
            i = False
        End Try
        Return i
    End Function

    Public Delegate Function CanExecuteDelegater(ByRef dat As RpaDataWrapper) As Integer

    Private _CanExecuteHandler As CanExecuteDelegater
    Public Property CanExecuteHandler As CanExecuteDelegater
        Get
            Return Me._CanExecuteHandler
        End Get
        Set(value As CanExecuteDelegater)
            Me._CanExecuteHandler = value
        End Set
    End Property


    Public Overloads Function Execute() As Integer Implements RpaCommandInterface.Execute
        Throw New NotImplementedException()
    End Function

    Public Overridable Overloads Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Dim i As Integer = -1
        Try
            Dim dlg As ExecuteDelegater = Me.ExecuteHandler
            If dlg Is Nothing Then
                Throw New NotImplementedException()
            Else
                i = dlg(dat)
            End If
        Catch ex As Exception
            Console.WriteLine($"({Me.GetType.Name}.ExecuteHandler) {ex.Message}")
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
