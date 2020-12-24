Public Class AddUtilityCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property CanExecute(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableParameterCount() As Integer()
        Get
            Return {1, 1}
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableProjectArchitectures As Integer()
        Get
            Return {
                RpaCodes.ProjectArchitecture.IntranetClientServer,
                RpaCodes.ProjectArchitecture.ClientServer,
                RpaCodes.ProjectArchitecture.StandAlone
            }
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableUserLevel() As Integer
        Get
            Return RpaCodes.ProjectUserLevel.User
        End Get
    End Property

    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        Dim uobj As RpaUtility
        Dim obj As Object
        Dim util As String
        If trn.Parameters.Count = 0 Then
            Console.WriteLine("ユーティリティが指定されていません")
            Return 1000
        End If
        util = trn.Parameters(0)
        obj = RpaCodes.RpaUtility(util)
        If obj Is Nothing Then
            Console.WriteLine($"指定のユーティリティ '{util}' は存在しません")
            Return 1000
        End If
        If rpa.SystemUtilities.ContainsKey(util) Then
            Console.WriteLine($"指定のユーティリティ '{util}' は既に追加されています")
            Return 1000
        End If
        uobj = New RpaUtility With {
            .UtilityName = util,
            .UtilityObject = obj
        }
        rpa.SystemUtilities.Add(util, uobj)
        Console.WriteLine($"指定のユーティリティ '{util}' を追加しました")
        Return 0
    End Function
End Class
