Public Class FileManagerModel
    Inherits ProjectBaseModel(Of FileManagerModel)
    'Public Overrides Sub MemberCheck()
    '    ' メンバのデフォルト値を設定したい場合、ここに記述
    '    If String.IsNullOrEmpty(Me.CurrentModelJson) Then
    '        Me.CurrentModelJson = Me.DefaultModelJson
    '    End If
    '    If String.IsNullOrEmpty(Me.CurrentInitViewModelJson) Then
    '        Me.CurrentInitViewModelJson = Me.DefaultInitViewModelJson
    '    End If
    '    If String.IsNullOrEmpty(Me.CurrentDBExplorerViewModelJson) Then
    '        Me.CurrentDBExplorerViewModelJson = Me.DefaultDBExplorerViewModelJson
    '    End If
    'End Sub

    Public Overrides Sub MemberCheck()
        ' メンバのデフォルト値を設定したい場合、ここに記述
        If String.IsNullOrEmpty(Me.CurrentModelJson) Then
            Me.CurrentModelJson = Me.DefaultModelJson
        End If
        If String.IsNullOrEmpty(Me.CurrentInitViewModelJson) Then
            Me.CurrentInitViewModelJson = Me.DefaultInitViewModelJson
        End If
        If String.IsNullOrEmpty(Me.CurrentDBExplorerViewModelJson) Then
            Me.CurrentDBExplorerViewModelJson = Me.DefaultDBExplorerViewModelJson
        End If
    End Sub

    Sub New()
        Call Me.MemberCheck()
    End Sub
End Class
