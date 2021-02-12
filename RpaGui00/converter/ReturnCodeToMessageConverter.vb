Public Class ReturnCodeToMessageConverter : Implements IValueConverter
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim sc As Integer = CType(value, Integer)
        Select Case sc
            Case -1
                Return $"コマンドはまだ実行されていません"
            Case 0
                Return $"コマンドは正常終了しました"

            ' CheckRobotCommandのチェックエラー
            Case 8100
                Return $"ロボットが起動条件を満たしていません"
            Case 8101
                Return $"プロジェクトが起動条件を満たしていません"

            ' RunRobotCommandのチェックエラー
            Case 8200
                Return $"ロボットが起動条件を満たしていません"
            Case 8201
                Return $"プロジェクトが起動条件を満たしていません"
            Case Else
                Return vbNullString
        End Select
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
