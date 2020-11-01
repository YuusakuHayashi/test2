Imports System.Globalization

Public Class DoubleGridLengthConverter : Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Return New GridLength(CType(value, Double))
        'Return (CType(value, GridLength)).Value
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        'Return New GridLength(CType(value, Double))
        Return (CType(value, GridLength)).Value
    End Function
End Class
