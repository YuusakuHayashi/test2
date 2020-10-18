Public Class DelegateAction
    Public Property CanExecuteHandler As Func(Of Object, Boolean)
    Public Property CanExecuteHandler2 As Func(Of Object, Boolean)
    Public Property CanExecuteHandler3 As Func(Of Object, Boolean)
    Public Property CanExecuteHandler4 As Func(Of Object, Boolean)
    Public Property CanExecuteHandler5 As Func(Of Object, Boolean)
    Public Property CanExecuteHandler6 As Func(Of Object, Boolean)
    Public Property CanExecuteHandler7 As Func(Of Object, Boolean)
    Public Property CanExecuteHandler8 As Func(Of Object, Boolean)
    Public Property CanExecuteHandler9 As Func(Of Object, Boolean)
    Public Property CanExecuteHandler10 As Func(Of Object, Boolean)
    Public Property ExecuteHandler As Action(Of Object)
    Public Property EvaluateHandler As Func(Of Object, Object)
    Public Property ExecuteIfFailedHandler As Action(Of Object)

    Public Sub Execute(ByVal parameter As Object)
        Dim e = Me.ExecuteHandler
        If e <> Nothing Then
            Call e(parameter)
        End If
    End Sub

    Public Function CanExecute(ByVal parameter As Object) As Boolean
        Dim ce = CanExecuteHandler
        Return IIf(ce = Nothing, False, ce(parameter))
    End Function

    Public Function ExecuteIfCan(ByVal parameter As Object) As Integer
        Dim i = -1
        Dim e = Me.ExecuteHandler
        Dim eif = Me.ExecuteIfFailedHandler
        Dim lce As New List(Of Func(Of Object, Boolean))

        lce.Add(Me.CanExecuteHandler)
        lce.Add(Me.CanExecuteHandler2)
        lce.Add(Me.CanExecuteHandler3)
        lce.Add(Me.CanExecuteHandler4)
        lce.Add(Me.CanExecuteHandler5)
        lce.Add(Me.CanExecuteHandler6)
        lce.Add(Me.CanExecuteHandler7)
        lce.Add(Me.CanExecuteHandler8)
        lce.Add(Me.CanExecuteHandler9)
        lce.Add(Me.CanExecuteHandler10)

        For Each ce In lce
            If ce <> Nothing Then
                i = 0
                If Not ce(parameter) Then
                    i = 1000
                    If eif <> Nothing Then
                        Call eif(parameter)
                        i = 100
                    End If
                    Exit For
                End If
            End If
        Next

        If i = 0 Then
            Call e(parameter)
        End If

        ExecuteIfCan = i
    End Function

    Public Function EvaluateIfCan(ByVal parameter As Object) As Object
        Dim i = -1
        Dim obj As Object : obj = i
        Dim e = Me.EvaluateHandler
        Dim eif = Me.ExecuteIfFailedHandler
        Dim lce As New List(Of Func(Of Object, Boolean))

        lce.Add(Me.CanExecuteHandler)
        lce.Add(Me.CanExecuteHandler2)
        lce.Add(Me.CanExecuteHandler3)
        lce.Add(Me.CanExecuteHandler4)
        lce.Add(Me.CanExecuteHandler5)
        lce.Add(Me.CanExecuteHandler6)
        lce.Add(Me.CanExecuteHandler7)
        lce.Add(Me.CanExecuteHandler8)
        lce.Add(Me.CanExecuteHandler9)
        lce.Add(Me.CanExecuteHandler10)

        For Each ce In lce
            If ce <> Nothing Then
                i = 0
                If Not ce(parameter) Then
                    i = 1000
                    If eif <> Nothing Then
                        Call eif(parameter)
                        i = 100
                    End If
                    Exit For
                End If
            End If
        Next

        If i = 0 Then
            obj = e(parameter)
        End If

        EvaluateIfCan = obj
    End Function

    Public Function CheckCanExecute(ByVal parameter As Object) As Integer
        Dim i = -1
        Dim lce As New List(Of Func(Of Object, Boolean))

        lce.Add(Me.CanExecuteHandler)
        lce.Add(Me.CanExecuteHandler2)
        lce.Add(Me.CanExecuteHandler3)
        lce.Add(Me.CanExecuteHandler4)
        lce.Add(Me.CanExecuteHandler5)
        lce.Add(Me.CanExecuteHandler6)
        lce.Add(Me.CanExecuteHandler7)
        lce.Add(Me.CanExecuteHandler8)
        lce.Add(Me.CanExecuteHandler9)
        lce.Add(Me.CanExecuteHandler10)

        For Each ce In lce
            i = 0
            If ce <> Nothing Then
                If Not ce(parameter) Then
                    i += 1
                    Exit For
                End If
            End If
        Next

        CheckCanExecute = i
    End Function
End Class
