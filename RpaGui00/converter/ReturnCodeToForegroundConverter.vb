Imports System.Drawing

Public Class ReturnCodeToForegroundConverter : Implements IValueConverter
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim sc As Integer = CType(value, Integer)
        Select Case sc
            Case -1
                Return Color.FromName("Black")
            Case 0
                Return Color.FromName("LimeGreen")

            ' CheckRobotCommandのチェックエラー
            Case 8100
                Return Color.FromName("Red")
            Case 8101
                Return Color.FromName("Red")

            ' RunRobotCommandのチェックエラー
            Case 8200
                Return Color.FromName("Red")
            Case 8201
                Return Color.FromName("Red")
            Case Else
                Return Color.FromName("Black")
        End Select
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
