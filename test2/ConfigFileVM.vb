Imports System.ComponentModel
Imports System.Windows.Forms

Public Class ConfigFileVM
    Inherits ViewModel

    Private _ConfigFileName As String
    Public Property ConfigFileName As String
        Get
            Return _ConfigFileName
        End Get
        Set(value As String)
            _ConfigFileName = value
        End Set
    End Property

    Private Sub ClickCommandExecute(ByVal parameter As Object)
        Dim ofd As New OpenFileDialog
        ofd.FileName = "default.html"
    End Sub

    Sub New()
        Me.ConfigFileName = "C:\Users\yuusaku.hayashi\test\ConfigFile.json"
    End Sub
End Class
