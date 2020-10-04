Public Class HistoryModel : Inherits BaseModel(Of HistoryModel)

    Private Const DELETE_OVER_LINES As String = "DeleteOverLines"
    Private Const RESET_HISTORIES As String = "ResetHistories"

    Private _Contents As String
    Public Property Contents As String
        Get
            Return Me._Contents
        End Get
        Set(value As String)
            Me._Contents = value
            RaisePropertyChanged("Contents")
        End Set
    End Property

    Private _BehavierWhenMaxLine As String
    Public Property BehavierWhenMaxLine As String
        Get
            Return Me._BehavierWhenMaxLine
        End Get
        Set(value As String)
            Me._BehavierWhenMaxLine = value
        End Set
    End Property

    Public Sub NewLine(ByVal txt As String)
        Me.Contents = txt
        Me._LineCounts = 1
    End Sub

    Public Sub AddLine(ByVal txt As String)
        If Me.Contents = vbNullString Then
            Me.Contents = txt
        Else
            Me.Contents &= vbCrLf & txt
        End If
        Me._LineCounts += 1
    End Sub

    Public Sub ClearContents()
        Me.Contents = ""
        Me._LineCounts = 0
    End Sub

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
            Dim dol As Action : dol = AddressOf _DeleteOverLines
            Dim reset As Action : reset = AddressOf _ResetBecauseOfMaxLine
            Me.__LineCounts = value
            If Me.__LineCounts > Me.MaxLine Then
                Select Case Me.BehavierWhenMaxLine
                    Case DELETE_OVER_LINES
                        dol()
                    Case RESET_HISTORIES
                        reset()
                    Case Else
                        reset()
                End Select
            End If
        End Set
    End Property

    Public Sub _ResetBecauseOfMaxLine()
        Me.NewLine("履歴が最大行を超えたため、クリアしました")
    End Sub

    Public Sub CheckLineCounts()
        Dim d As Double
        d = (Me.Contents.Length - Me.Contents.Replace(vbCrLf, "").Length) / 2
        Me._LineCounts = Math.Truncate(d)
    End Sub

    Private Sub _DeleteOverLines()
        Dim idx As Long
        idx = Me.Contents.IndexOf(vbCrLf)
        Me.Contents = Me.Contents.Remove(0, Convert.ToInt32(idx) + 2)
        Me._LineCounts -= 1
    End Sub

    'Public Overrides Sub MemberCheck()
    '    If Me.MaxLine = Nothing Or 0 Then
    '        Me.MaxLine = 100
    '    End If
    '    If String.IsNullOrEmpty(Me.Contents) Then
    '        Me.NewLine("Historyは初期化されました")
    '    End If
    '    If String.IsNullOrEmpty(Me.BehavierWhenMaxLine) Then
    '        Me.BehavierWhenMaxLine = RESET_HISTORIES
    '    End If
    'End Sub

    'Sub New()
    '    Call Me.MemberCheck()
    'End Sub
End Class
