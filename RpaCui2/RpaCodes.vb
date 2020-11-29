Module RpaCodes

    Public ReadOnly Property RpaObject(ByVal code As String) As Object
        Get
            Select Case code
                Case "rpa01"
                    Return New Rpa01
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property
End Module
