Imports System.IO

Module ProjectModule
    ' この関数はプロジェクトディレクトリの存在チェックを行います
    Public Function CheckProjectDirectory(ByVal elm As ProjectInfoModel) As Boolean
        Dim b As Boolean : b = False
        If Directory.Exists(elm.DirectoryName) Then
            If File.Exists(elm.IniFileName) Then
                b = True
            End If
        End If
        CheckProjectDirectory = b
    End Function
End Module
