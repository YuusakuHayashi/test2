Public Class DensityToGreenConverter : Implements IValueConverter

    Private Shared _GreenColors As List(Of Brush)
    Private Shared ReadOnly Property GreenColors As List(Of Brush)
        Get
            If _GreenColors Is Nothing Then
                Dim c As New BrushConverter
                _GreenColors = New List(Of Brush)
                _GreenColors.Add(c.ConvertFrom("#FFEBEDF0"))
                _GreenColors.Add(c.ConvertFrom("#FF9BE9A8"))
                _GreenColors.Add(c.ConvertFrom("#FF40C463"))
                _GreenColors.Add(c.ConvertFrom("#FF30A14E"))
                _GreenColors.Add(c.ConvertFrom("#FF216E39"))
            End If
            Return _GreenColors
        End Get
    End Property

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim count As Integer = CType(value, Integer)
        If count > (GreenColors.Count - 1) Then
            count = (GreenColors.Count - 1)
        End If
        Return GreenColors(count)
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
