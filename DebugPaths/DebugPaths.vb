Imports System.IO

Public Class DebugPaths : Inherits RpaCui.JsonHandler(Of DebugPaths)
    Public PathDictionary As Dictionary(Of String, String)

    Public ReadOnly Property DebugPathsFileName As String
        Get
            Dim [dir] As String = $"{RpaCui.SystemDirectory}\debugpaths"
            If Not File.Exists([dir]) Then
                Me.Save([dir], Me)
            End If
            Return [dir]
        End Get
    End Property
End Class
