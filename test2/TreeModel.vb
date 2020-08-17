Imports System.Collections.ObjectModel

Public Class TreeModel
    Inherits BaseModel
    Implements IEnumerable(Of TreeModel)

    '
    Private _RealName As String
    Public Property RealName As String
        Get
            Return Me._RealName
        End Get
        Set(value As String)
            Me._RealName = value
        End Set
    End Property

    '
    Public ReadOnly Property Children As ObservableCollection(Of TreeModel)
        Get
            Return New ObservableCollection(Of TreeModel)
        End Get
    End Property

    '
    Private _IsSelected As Boolean
    Public Property IsSelected As Boolean
        Get
            Return Me._IsSelected
        End Get
        Set(value As Boolean)
            Me._IsSelected = value
        End Set
    End Property

    '
    Private _IsExpanded As Boolean
    Public Property IsExpanded As Boolean
        Get
            Return Me._IsExpanded
        End Get
        Set(value As Boolean)
            Me._IsExpanded = value
        End Set
    End Property

    '
    Private _IsChecked As Boolean
    Public Property IsChecked As Boolean
        Get
            Return Me._IsChecked
        End Get
        Set(value As Boolean)
            Me._IsChecked = value
        End Set
    End Property

    Friend Parent As TreeModel

    '
    Public Sub Add(c As TreeModel)
        c.Parent = Me
        Me.Children.Add(c)
    End Sub

    '
    Public Sub Remove(c As TreeModel)
        Me.Children.Remove(c)
        c.Parent = Nothing
    End Sub

    '
    Sub New(rn As String)
        Me.RealName = rn
    End Sub

    Public Function GetEnumerator() As IEnumerator(Of TreeModel) Implements IEnumerable(Of TreeModel).GetEnumerator
        Throw New NotImplementedException()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Throw New NotImplementedException()
    End Function
End Class
