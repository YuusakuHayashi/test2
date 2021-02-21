Imports System.Globalization

Public Class DoubleToMonthStringsConverter : Implements IValueConverter
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim dbl As Double = CType(value, Double)
        Select Case dbl
            Case 0.0 : Return vbNullString
            Case 0.5 : Return "-/Jan."
            Case 1.0 : Return "Jan."
            Case 1.5 : Return "-/Feb."
            Case 2.0 : Return "Feb."
            Case 2.5 : Return "-/Mar."
            Case 3.0 : Return "Mar."
            Case 3.5 : Return "-/Apr."
            Case 4.0 : Return "Apr."
            Case 4.5 : Return "-/May"
            Case 5.0 : Return "May"
            Case 5.5 : Return "-/Jun."
            Case 6.0 : Return "Jun."
            Case 6.5 : Return "-/Jul."
            Case 7.0 : Return "Jul."
            Case 7.5 : Return "-/Aug."
            Case 8.0 : Return "Aug."
            Case 8.5 : Return "-/Sep."
            Case 9.0 : Return "Sep."
            Case 9.5 : Return "-/Oct."
            Case 10.0 : Return "Oct."
            Case 10.5 : Return "-/Nov."
            Case 11.0 : Return "Nov."
            Case 11.5 : Return "-/Dec."
            Case 12.0 : Return "Dec."
            Case 12.5 : Return "-/Jan."
            Case Else : Return vbNullString
        End Select
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
