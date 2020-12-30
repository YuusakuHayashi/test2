Imports System.Reflection

Public Class ShowProjectPropertiesCommand : Inherits RpaCommandBase
    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        ' 初回時にプロパティ一覧を表示
        Dim pname As String = vbNullString
        Dim ptype As String = vbNullString
        Dim getset As String = vbNullString
        Dim pnamelen As Integer = 25
        Dim ptypelen As Integer = 15

        Dim props As PropertyInfo() = rpa.GetType.GetProperties()
        For Each prop In props
            ' Property Name
            pname = prop.Name
            If pname.Length > pnamelen Then
                pname = Strings.Left(pname, (pnamelen - 3)) & "..."
            ElseIf pname.Length < pnamelen Then
                pname = pname & Strings.StrDup((pnamelen - pname.Length), " "c)
            Else
                'Nothing To Do
            End If

            If (prop.CanRead) And (prop.CanWrite) Then
                getset = "Get/Set"
            ElseIf (prop.CanRead) And (Not prop.CanWrite) Then
                getset = "  Get  "
            ElseIf (Not prop.CanRead) And (prop.CanWrite) Then
                getset = "  Set  "
            Else
                getset = "-------"
            End If

            ' Property Type
            ptype = prop.PropertyType.Name
            If ptype.Length > ptypelen Then
                ptype = Strings.Left(ptype, (ptypelen - 3)) & "..."
            ElseIf ptype.Length < ptypelen Then
                ptype = ptype & Strings.StrDup((ptypelen - ptype.Length), " "c)
            Else
                'Nothing To Do
            End If
            Console.Write($"{pname} {getset} {ptype} ")
            Try
                Console.WriteLine($"{prop.GetValue(rpa).ToString}")
            Catch e As Exception
                Console.WriteLine($"{e.Message}")
            End Try
        Next
        Console.WriteLine()
        Return 0
    End Function
End Class
