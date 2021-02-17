
Public Class StatusCodeToColorConverter : Implements IValueConverter
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim sc As Integer = CType(value, Integer)
        Select Case sc
            Case 0
                Return Media.Color.FromArgb(&HFF, 33, 22, 16)
            Case 1
                Return Media.Color.FromArgb(&HFF, 230, 82, 38)
            Case Else
                Return Media.Color.FromArgb(&HFF, 33, 22, 16)
        End Select
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
