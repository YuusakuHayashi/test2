Imports System.Globalization

Public Class OverTimeToColorConverter : Implements IValueConverter
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim flag As Boolean = CType(value, Boolean)
        If flag Then
            Return Colors.Yellow
        Else
            Return Colors.Green
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
