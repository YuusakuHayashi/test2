Public Class ExecuteCountToGreenConverter : Implements IValueConverter
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim count As Integer = CType(value, Integer)
        Dim green As Integer = 220 - (count * 10)
        If green < 0 Then
            green = 0
        End If
        If count > 0 Then
            Return Media.Color.FromArgb(&HFF, 0, green, 0)
        Else
            Return Colors.WhiteSmoke
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
