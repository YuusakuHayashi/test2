Imports System.Runtime.InteropServices
Imports System.IO
Imports Rpa00

' Note : 
'     いつごろからか分からないが、マクロを呼び出ししようとすると、例外を出すようになった(0x8002005 (DISP_E_TYPEMISMATCH))
'     根本的な原因はわからないが、InvokeMacro(ByVal method As String, ByRef args() As Object) などは、
'     パラメタの要求として、オブジェクト型の配列を現状要求するようにしている。 
'     だが、配列の中身（各要素）がオブジェクト型ではない場合、この例外が発生することが分かった。
'     ただ、このクラス内にさらに内部クラス宣言している UpdateMacroCommand から InvokeMacro を呼び出す場合、
'     あからさまに文字列型の要素をセットしているが、例外は発生しない。
'     可能性として、内部クラスの場合、親クラスのメソッドの定義を知っているが、dllを外部から呼び出す場合は、
'     Object のため、メソッド定義が分からず、実行時例外が発生しているのか？・・・

Public Class RpaMacroUtility : Inherits RpaUtilityBase
    Public Overrides ReadOnly Property CommandHandler(dat As Object) As Object
        Get
            Dim cmd As Object
            Select Case dat.Transaction.MainCommand
                Case "update" : cmd = New UpdateMacroCommand(Me)
                Case Else : cmd = Nothing
            End Select
            Return cmd
        End Get
    End Property

    Public ReadOnly Property MacroUtilityDirectory As String
        Get
            Dim d = $"{CommonProject.SystemDirectory}\MacroUtility"
            If Not Directory.Exists(d) Then
                Directory.CreateDirectory(d)
            End If
            Return d
        End Get
    End Property

    Public ReadOnly Property MacroFileName As String
        Get
            Dim f = $"{Me.MacroUtilityDirectory}\macro.xlsm"
            Return f
        End Get
    End Property

    Private Function CallMacro(ByVal method As String, ByRef args() As Object) As Object
        Dim obj As Object
        Dim exapp, exbooks, exbook
        Dim macrofile As String = vbNullString

        Try
            exapp = CreateObject("Excel.Application")
            Try
                exbooks = exapp.Workbooks
                Try
                    exbook = exbooks.Open(Me.MacroFileName, 0)
                    macrofile = Path.GetFileName(Me.MacroFileName)
                    obj = exapp.Run($"{macrofile}!{method}", args)
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                    Console.WriteLine($"Workbookオブジェクトの取得、")
                    Console.WriteLine($"または、'{macrofile}!{method}' の実行に失敗しました")
                    obj = Nothing
                Finally
                    If exbook IsNot Nothing Then
                        exbook.Close()
                    End If
                    Marshal.ReleaseComObject(exbook)
                End Try
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Console.WriteLine($"Workbooksオブジェクトの取得に失敗しました")
                obj = Nothing
            Finally
                Marshal.ReleaseComObject(exbooks)
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine($"Excel.Applicationオブジェクトの取得に失敗しました")
            obj = Nothing
        Finally
            If exapp IsNot Nothing Then
                exapp.Quit()
            End If
            Marshal.ReleaseComObject(exapp)
        End Try
        Return obj
    End Function

    Private Function ConvertObjectArgs(ByRef args() As Object) As Object()
        Dim objs As New List(Of Object)
        For Each arg In args
            objs.Add(CType(arg, Object))
        Next
        Return (objs.ToArray())
    End Function

    Public Function InvokeMacro(ByVal method As String, ByRef args() As Object) As Integer
        Dim newargs() = ConvertObjectArgs(args)
        Dim obj = Me.CallMacro(method, newargs)
        Return IIf((obj Is Nothing), 0, 1000)
    End Function

    Public Function InvokeMacroFunction(ByVal method As String, ByRef args() As Object) As Object
        Dim newargs() = ConvertObjectArgs(args)
        Dim obj = Me.CallMacro(method, newargs)
        Return obj
    End Function

    Private Class UpdateMacroCommand : Inherits RpaUtilityCommandBase
        Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
            Get
                Return {1, 999}
            End Get
        End Property

        Public Overrides Function CanExecute(ByRef dat As Object) As Boolean
            If Not File.Exists(Me.Parent.MacroFileName) Then
                Console.WriteLine($"マクロファイル '{Me.Parent.MacroFileName}' が存在しません")
                Console.WriteLine($"開発者から入手してください")
                Console.WriteLine()
                Return False
            End If
            Return True
        End Function

        Private Parent As RpaMacroUtility
        Public Overrides Function Execute(ByRef dat As Object) As Integer
            For Each p In dat.Transaction.Parameters
                Dim bas As String = $"{Parent.MacroUtilityDirectory}\{p}"
                If File.Exists(bas) Then
                    Console.WriteLine($"指定マクロ '{p}' をインストールします")
                    Call Parent.InvokeMacro("MacroImporter.Main", {bas})
                    Console.WriteLine()
                Else
                    Console.WriteLine($"指定マクロ '{p}' は存在しません")
                    Console.WriteLine()
                End If
            Next
            Return 0
        End Function

        Sub New(p As RpaMacroUtility)
            Me.Parent = p
        End Sub
    End Class
End Class
