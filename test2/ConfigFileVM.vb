Imports System.ComponentModel
Imports System.Windows.Forms

Public Class ConfigFileVM
    Inherits ViewModel

    Private Const _PLEASE_INPUT = "Please Input Config File Name"

    Private _ConfigFileName As String
    Public Property ConfigFileName As String
        Get
            If _ConfigFileName <> vbNullString Then
                Return _ConfigFileName
            Else
                Return _PLEASE_INPUT
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
        Me.ConfigFileName = _ConfigFileName
    End Sub
End Class
