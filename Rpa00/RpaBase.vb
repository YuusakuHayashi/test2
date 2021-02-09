Imports Newtonsoft.Json

Public MustInherit Class RpaBase(Of T As {New}) : Inherits Rpa00.JsonHandler(Of T)
    Public MustOverride Function SetupProjectObject(ByVal project As String) As Object
    'Public MustOverride Function Execute(ByRef dat As Object) As Integer
    'Public MustOverride Function CanExecute(ByRef dat As Object) As Boolean

    Public Delegate Function ExecuteDelegater(ByRef dat As Object) As Integer
    Private _ExecuteHandler As ExecuteDelegater
    <JsonIgnore>
    Public Property ExecuteHandler As ExecuteDelegater
        Get
            Return Me._ExecuteHandler
        End Get
        Set(value As ExecuteDelegater)
            Me._ExecuteHandler = value
        End Set
    End Property

    Public Function Execute(ByRef dat As Object) As Integer
        Dim i As Integer = False
        Try
            Dim dlg As ExecuteDelegater = Me.ExecuteHandler
            If dlg Is Nothing Then
                Throw New NotImplementedException
            Else
                i = dlg(dat)
            End If
        Catch ex As Exception
            Console.WriteLine($"({Me.GetType.Name}.Execute) {ex.Message}")
            Console.WriteLine()
            i = -1
        End Try
        Return i
    End Function

    Public Delegate Function CanExecuteDelegater(ByRef dat As Object) As Boolean
    Private _CanExecuteHandler As CanExecuteDelegater
    <JsonIgnore>
    Public Property CanExecuteHandler As CanExecuteDelegater
        Get
            Return Me._CanExecuteHandler
        End Get
        Set(value As CanExecuteDelegater)
            Me._CanExecuteHandler = value
        End Set
    End Property

    Public Function CanExecute(ByRef dat As Object) As Boolean
        Dim b As Boolean = False
        Try
            Dim dlg As CanExecuteDelegater = Me.CanExecuteHandler
            If dlg Is Nothing Then
                Throw New NotImplementedException
            Else
                b = dlg(dat)
            End If
        Catch ex As Exception
            Console.WriteLine($"({Me.GetType.Name}.CanExecute) {ex.Message}")
            Console.WriteLine()
            b = False
        End Try
        Return b
    End Function

    'Public Function CanExecute(dat As Object) As Boolean
    '    Dim b As Boolean = False
    '    Try
    '        Dim dlg As CanAsyncExecuteDelegater = Me.CanAsyncExecuteHandler
    '        If dlg Is Nothing Then
    '            Throw New NotImplementedException
    '        Else
    '            Dim t As Task(Of Boolean) = Task.Run(
    '                Function()
    '                    Return dlg(dat)
    '                End Function
    '            )
    '            b = t.Result
    '        End If
    '    Catch ex As Exception
    '        Console.WriteLine($"({Me.GetType.Name}.CanExecute) {ex.Message}")
    '        Console.WriteLine()
    '        b = False
    '    End Try
    '    Return b
    'End Function

    Public Delegate Function CanAsyncExecuteDelegater(dat As Object) As Task(Of Boolean)
    Private _CanAsyncExecuteHandler As CanAsyncExecuteDelegater
    <JsonIgnore>
    Public Property CanAsyncExecuteHandler As CanAsyncExecuteDelegater
        Get
            Return Me._CanAsyncExecuteHandler
        End Get
        Set(value As CanAsyncExecuteDelegater)
            Me._CanAsyncExecuteHandler = value
        End Set
    End Property

    <JsonIgnore>
    Public Data As Object
End Class
