Imports Rpa00

Public Class AddUtilityCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim util As String = dat.System.CommandData.Parameters(0)
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

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecutableParameterCount = {1, 1}
    End Sub
End Class
