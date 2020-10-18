Public Class HistoryModel : Inherits BaseModel(Of Object)
    Private Const DELETE_OVERLINES As String = "DeleteOverLines"
    Private Const RESET_HISTORIES As String = "ResetHistories"

    Public ExecuteWhenMaxLineMethod As String
    Private _ExecuteWhenMaxLine As Action

    Public ReadOnly Property DisplayContents
        Get
            Return Me._Contents
        End Get
    End Property

    Private __Contents As String
    Public Property _Contents As String
        Get
            Return Me.__Contents
        End Get
        Set(value As String)
            Me.__Contents = value
            RaisePropertyChanged("_Contents")
            RaisePropertyChanged("DisplayContents")
        End Set
    End Property

    Public Sub NewLine(ByVal txt As String)
        Me._Contents = txt & vbCrLf
        Me._LineCounts = 1
    End Sub

    Public Sub AddLine(ByVal txt As String)
        Me._Contents &= txt & vbCrLf
        Me._LineCounts += 1
    End Sub

    'Public Sub Clear_Contents()
    '    Me._Contents = ""
    '    Me._LineCounts = 0
    'End Sub

    Private _MaxLine As Integer
    Public Property MaxLine As Integer
        Get
            Return Me._MaxLine
        End Get
        Set(value As Integer)
            Me._MaxLine = value
        End Set
    End Property

    Private __LineCounts As Integer
    Private Property _LineCounts As Integer
        Get
            Return Me.__LineCounts
        End Get
        Set(value As Integer)
            Me.__LineCounts = value
            If Me.__LineCounts > Me.MaxLine Then
                Call Me._ExecuteWhenMaxLine()
            End If
        End Set
    End Property

    Private Sub _GetLineCounts()
        Dim d = (Me._Contents.Length - Me._Contents.Replace(vbCrLf, "").Length) / 2
        Me._LineCounts = Math.Truncate(d)
    End Sub

    Private Sub _ResetHistories()
        Call Me.NewLine("履歴が最終行を超え、クリアされました")
    End Sub

    Private Sub _DeleteOverLines()
        Dim idx = Me._Contents.IndexOf(vbCrLf)
        Me._Contents = Me.__Contents.Remove(0, Convert.ToInt32(idx) + 2)
        Call Me._GetLineCounts()
    End Sub

    Sub New()
        If MaxLine = 0 Then
            Me.MaxLine = 100
        End If

        If Not String.IsNullOrEmpty(Me._Contents) Then
            Call Me._GetLineCounts()
        End If

        Select Case ExecuteWhenMaxLineMethod
            Case DELETE_OVERLINES
                Me._ExecuteWhenMaxLine = AddressOf Me._DeleteOverLines
            Case RESET_HISTORIES
                Me._ExecuteWhenMaxLine = AddressOf Me._ResetHistories
            Case Else
                Me._ExecuteWhenMaxLine = AddressOf Me._DeleteOverLines
        End Select
    End Sub
End Class
