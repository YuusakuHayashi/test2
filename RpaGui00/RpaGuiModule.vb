Public Module RpaGuiModule
    Public Function CreateIcon(ByVal filename As String) As BitmapImage
        Dim bi = New BitmapImage
        bi.BeginInit()
        bi.UriSource = New Uri(filename, UriKind.Absolute)
        bi.EndInit()
        Return bi
    End Function
End Module
