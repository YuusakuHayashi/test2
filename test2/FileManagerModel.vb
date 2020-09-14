Imports System.IO
Public Class FileManagerModel
    Inherits ProjectModel(Of FileManagerModel)

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

    ' このモデルは初期起動時に、各種モデルの設定ファイルをロードする

    Public Sub MemberCheck()
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


    ' 各モデルファイルのファイル名をセット
    Public Sub CurrentJsonSet()
        Dim m As Model

        ' Model
        m = Me.ModelLoad(Of Model)(Me.CurrentModelJson)
        If m Is Nothing Then
            m = New Model
        End If
        If String.IsNullOrEmpty(m.CurrentModelJson) Then
            m.CurrentModelJson = Me.CurrentModelJson
        End If
        If Not File.Exists(m.CurrentModelJson) Then
            m.CurrentModelJson = Me.CurrentModelJson
        Else
            Me.CurrentModelJson = m.CurrentModelJson
        End If
        m.ModelSave(Of Model)(m.CurrentModelJson, m)
    End Sub

    Sub New()
        Me.MemberCheck()
        Me.CurrentJsonSet()
    End Sub
End Class
