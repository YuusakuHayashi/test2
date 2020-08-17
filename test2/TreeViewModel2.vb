Imports System.ComponentModel

Public Class TreeViewModel2 : Inherits TreeViewItem

    Private _Expand As Boolean

    Private __Origin As TreeViewOrigin
    Private Property _Origin As TreeViewOrigin
        Get
            Return __Origin
        End Get
        Set(value As TreeViewOrigin)
            __Origin = value
        End Set
    End Property


    Private _RealName As String
    Public Property RealName As String
        Get
            Return Me._RealName
        End Get
        Set(value As String)
            Me._RealName = value
        End Set
    End Property


    Private _SelectItem As TreeViewModel2
    Public Property SelectItem As TreeViewModel2
        Get
            Return _SelectItem
        End Get
        Set(value As TreeViewModel2)
            _SelectItem = value
        End Set
    End Property


    '展開時に子要素をすべてインスタンス化
    Private Sub _ExpandedBehavior(ByVal sender As Object, ByVal e As RoutedEventArgs)
        If Not Me._Expand Then
            Me.Items.Clear()
            For Each c In Me._Origin.Children
                Me.Items.Add(New TreeViewModel2(c))
            Next
            Me._Expand = True
        End If
    End Sub


    '
    Private Sub _SelectedBehavior(ByVal sender As Object, ByVal e As RoutedEventArgs)
        If Me.IsSelected Then
            Me.SelectItem = Me
        Else
            Me.SelectItem = e.Source
        End If
    End Sub


    Sub New(ByRef tvo As TreeViewOrigin)
        Me._Expand = False

        Me._Origin = tvo
        Me.RealName = tvo.RealName

        If tvo.Children IsNot Nothing Then
            '空のアイテムを追加しておくことで、展開の＞が表示される
            Me.Items.Add(New TreeViewItem())
            AddHandler Me.Expanded, AddressOf Me._ExpandedBehavior
        End If

        AddHandler Me.Selected, AddressOf Me._SelectedBehavior
    End Sub
End Class
