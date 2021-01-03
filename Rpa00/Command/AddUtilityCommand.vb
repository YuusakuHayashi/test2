Imports Rpa00

Public Class AddUtilityCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
        Get
            Return {1, 1}
        End Get
    End Property

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer

        Dim util As String = dat.Transaction.Parameters(0)
        Dim uobj As Object = RpaCodes.RpaUtilityObject(util)

        If uobj Is Nothing Then
            Console.WriteLine($"指定のユーティリティ '{util}' は存在しません")
            Console.WriteLine()
            Return 1000
        End If
        If dat.Project.SystemUtilities.ContainsKey(util) Then
            Console.WriteLine($"指定のユーティリティ '{util}' は既に追加されています")
            Console.WriteLine()
            Return 1000
        End If

        dat.Project.AddUtility(util)

        Console.WriteLine($"指定のユーティリティ '{util}' を追加しました")
        Console.WriteLine()
        Return 0
    End Function
End Class
