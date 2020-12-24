Imports System.IO

Public Class LoadCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecutableProjectArchitectures As Integer()
        Get
            Return {
                RpaCodes.ProjectArchitecture.ClientServer,
                RpaCodes.ProjectArchitecture.IntranetClientServer,
                RpaCodes.ProjectArchitecture.StandAlone
            }
        End Get
    End Property

    Public Overrides ReadOnly Property CanExecute(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableParameterCount() As Integer()
        Get
            Return {0, 0}
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableUserLevel() As Integer
        Get
            Return RpaCodes.ProjectUserLevel.User
        End Get
    End Property

    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        Dim [load] = rpa.Load(rpa.SystemJsonFileName)
        Dim i = 0
        If [load] Is Nothing Then
            Console.WriteLine($"JsonFile '{rpa.SystemJsonFileName}' のロードに失敗しました。")
            Return 1000
        End If
        'If Not Directory.Exists([load].RootDirectory) Then
        '    Console.WriteLine($"RootDirectory '{[load].RootDirectory}' がありません")
        '    Console.WriteLine($"ファイル '{[load].SystemJsonFileName}' の 'RootDirectory' に任意のパスを書いてください")
        '    Console.WriteLine("ファイルを保存した後、アプリケーションを再起動してください")
        '    Console.ReadLine()
        '    trn.ExitFlag = True
        '    Return 1000
        'End If
        'If Not Directory.Exists([load].MyDirectory) Then
        '    Console.WriteLine($"MyDirectory '{[load].MyDirectory}' がありません")
        '    Console.WriteLine($"ファイル '{[load].SystemJsonFileName}' の 'MyDirectory' に任意のパスを書いてください")
        '    Console.WriteLine("ファイルを保存した後、アプリケーションを再起動してください")
        '    Console.ReadLine()
        '    trn.ExitFlag = True
        '    Return 1000
        'End If
        rpa = [load]
        Console.WriteLine($"JsonFile '{rpa.SystemJsonFileName}' をロードしました。")
        Return 0
    End Function
End Class
