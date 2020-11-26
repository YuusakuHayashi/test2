Imports System.Globalization

Public Class FileTypeIconConverter : Implements IValueConverter
    Private ReadOnly Property _FileTypeIconName(ByVal ft As String) As String
        Get
            Select Case ft
                Case "File"
                    Return AppDirectoryModel.AppImageDirectory & "\file.png"
                Case "Folder"
                    Return AppDirectoryModel.AppImageDirectory & "\folder.png"
                Case Else
                    Return AppDirectoryModel.AppImageDirectory & "\NoIcon.png"
            End Select
        End Get
    End Property

    Private Shared __FileTypeIconDictionary As Dictionary(Of String, BitmapImage)
    Private Property _FileTypeIconDictionary As Dictionary(Of String, BitmapImage)
        Get
            If FileTypeIconConverter.__FileTypeIconDictionary Is Nothing Then
                FileTypeIconConverter.__FileTypeIconDictionary = New Dictionary(Of String, BitmapImage)
            End If
            Return FileTypeIconConverter.__FileTypeIconDictionary
        End Get
        Set(value As Dictionary(Of String, BitmapImage))
            FileTypeIconConverter.__FileTypeIconDictionary = value
        End Set
    End Property

    Private Function _GetFileTypeIcon(ByVal [key] As String) As BitmapImage
        If Not Me._FileTypeIconDictionary.ContainsKey([key]) Then
            Me._FileTypeIconDictionary([key]) = _AssignFileTypeIcon([key])
        End If
        Return Me._FileTypeIconDictionary([key])
    End Function

    Private Function _AssignFileTypeIcon(ByVal ft As String) As BitmapImage
        Dim bi = New BitmapImage
        bi.BeginInit()
        bi.UriSource = New Uri((Me._FileTypeIconName(ft)), UriKind.Absolute)
        bi.EndInit()
        Return bi
    End Function

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Return Me._GetFileTypeIcon(value)
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
