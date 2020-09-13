Imports System.ComponentModel
Imports System.Windows.Forms

Public Class ConfigFileVM
    Inherits BaseViewModel

    Public Const PLEASE_INPUT = "Please Input Config File Name"

    Private _ConfigFileName As String
    Public Property ConfigFileName As String
        Get
            If String.IsNullOrEmpty(_ConfigFileName) Then
                Return PLEASE_INPUT
            Else
                Return _ConfigFileName
            End If
        End Get
        Set(value As String)
            _ConfigFileName = value
            RaisePropertyChanged("ConfigFileName")
        End Set
    End Property

    Private Sub ClickCommandExecute(ByVal parameter As Object)
        Dim ofd As New OpenFileDialog
        ofd.FileName = "default.html"
    End Sub

    Sub New()

    End Sub
End Class
