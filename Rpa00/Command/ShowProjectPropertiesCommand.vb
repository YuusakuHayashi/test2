﻿Imports System.Reflection


' SetupProjectCommand中に実行可能
Public Class ShowProjectPropertiesCommand : Inherits RpaCommandBase
    'Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
    '    If dat.Project Is Nothing Then
    '        Console.WriteLine($"現在のプロジェクトがありません")
    '        Return False
    '    End If
    '    Return True
    'End Function

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        ' 初回時にプロパティ一覧を表示
        Dim pnamelen As Integer = 25
        Dim ptypelen As Integer = 15
        Console.WriteLine()
        Console.WriteLine($"Property Name             | Get/Set | Property Type   | Value")
        Console.WriteLine($"--------------------------+---------+-----------------+------")

        Dim props As PropertyInfo() = dat.Project.GetType.GetProperties()
        For Each prop In props
            ' Property Name
            Dim pname As String = prop.Name
            If pname.Length > pnamelen Then
                pname = Strings.Left(pname, (pnamelen - 3)) & "..."
            ElseIf pname.Length < pnamelen Then
                pname = pname & Strings.StrDup((pnamelen - pname.Length), " "c)
            Else
                'Nothing To Do
            End If

            Dim getset As String = vbNullString
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
            Dim ptype As String = prop.PropertyType.Name
            If ptype.Length > ptypelen Then
                ptype = Strings.Left(ptype, (ptypelen - 3)) & "..."
            ElseIf ptype.Length < ptypelen Then
                ptype = ptype & Strings.StrDup((ptypelen - ptype.Length), " "c)
            Else
                'Nothing To Do
            End If

            Console.Write($"{pname} | {getset} | {ptype} | ")

            Try
                Console.WriteLine($"{prop.GetValue(dat.Project).ToString}")
            Catch e As Exception
                Console.WriteLine($"{e.Message}")
            End Try
        Next
        Console.WriteLine()
        Return 0
    End Function
End Class
