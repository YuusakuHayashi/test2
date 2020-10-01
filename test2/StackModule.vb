Module StackModule
    Public Function Push(Of ColT As {IList, New}, T)(ByVal elm As T, ByRef [old] As ColT)
        Dim [new] As New ColT

        [new].Add(elm)

        For Each e In [old]
            [new].Add(e)
        Next

        Push = [new]
    End Function

    Public Sub Peek()
    End Sub
End Module
