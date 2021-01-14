Imports System.Reflection

Public Class ChangeInitializerPropertyCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
        Get
            Return {2, 2}
        End Get
    End Property

    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        Dim pname As String = dat.Transaction.Parameters(0)
        Dim pi As PropertyInfo = dat.Initializer.GetType().GetProperty(pname)
        If pi Is Nothing Then
            Console.WriteLine($"プロパティ '{pname}' は存在しません")
            Console.WriteLine()
            Return False
        End If
        If Not pi.CanWrite Then
            Console.WriteLine($"プロパティ '{pname}' は変更不可です")
            Console.WriteLine()
            Return False
        End If
        Return True
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim pname As String = dat.Transaction.Parameters(0)
        Dim ptext As String = dat.Transaction.Parameters(1)
        Dim pi As PropertyInfo = dat.Initializer.GetType().GetProperty(pname)

        Dim pvalue As Object
        Select Case pi.PropertyType
            Case (New Integer).GetType             ' 数値
                pvalue = Integer.Parse(ptext)
            Case "String".GetType                  ' 文字列
                pvalue = ptext
            Case True.GetType                      ' 論理値
                pvalue = Boolean.Parse(ptext)
            Case Else
                Console.WriteLine($"プロパティ '{pname}' は変更不可です")
                Console.WriteLine()
                pvalue = Nothing
        End Select

        If pvalue Is Nothing Then
            Console.WriteLine($"プロパティ '{pname}' を 値' {ptext}' に変更できませんでした")
            Console.WriteLine()
            Return 1000
        End If

        pi.SetValue(dat.Initializer, pvalue)
        Console.WriteLine($"プロパティ '{pname}' を 値' {ptext}' に変更しました")
        Console.WriteLine()

        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
    End Sub
End Class
