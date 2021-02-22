Public Class DensityToGreenConverter : Implements IValueConverter
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim count As Integer = CType(value, Integer)
        If count < 1 Then
            Return Colors.WhiteSmoke
        ElseIf count < 2 Then
            Return Color.FromArgb(&HFF, 204, 207, 154)
        ElseIf count < 3 Then
            Return Color.FromArgb(&HFF, 211, 225, 115)
        ElseIf count < 4 Then
            Return Colors.SeaGreen
        ElseIf count < 5 Then
            Return Colors.Green
        Else
            Return Color.FromArgb(&HFF, 0, 81, 51)
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
