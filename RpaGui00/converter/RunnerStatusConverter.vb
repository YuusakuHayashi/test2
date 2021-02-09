Public Class RunnerStatusConverter : Implements IValueConverter
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim sc As Integer = CType(value, Integer)
        Select Case sc
            Case 0
                Return $"初期状態です"
            Case 1
                Return $"実行中です"
            Case 2
                Return $"実行完了しました"
            Case Else
                Return $"Error StatusCode='{sc}'"
        End Select
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
