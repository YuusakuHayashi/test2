Module StackModule
    Public Function Push(Of ColT As {IList, New}, T)(ByVal elm As T, 
                                                     ByVal [old] As ColT, 
                                                     ByVal max As Integer)
        Dim [new] As New ColT

        [new].Add(elm)

        For Each e In [old]
            [new].Add(e)
            If [new].Count >= max Then
                Exit For
            End If
        Next

        Push = [new]
    End Function

    Public Sub Peek()
    End Sub
End Module
