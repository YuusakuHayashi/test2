Imports System.Globalization
Imports System.Drawing

Public Class RotateAngleConverter : Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim s = CType(value, Single)
        Select Case value
            Case "Bottom"
                Return Dock.Bottom
            Case "Left"
                Return Dock.Left
            Case "Right"
                Return Dock.Right
            Case "Top"
                Return Dock.Top
            Case Else
                Return Dock.Top
        End Select
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
