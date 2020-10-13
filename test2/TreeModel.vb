'Imports System.Collections.ObjectModel

'Public Class TreeModel
'    Inherits BaseModel(Of TreeModel)
'    Implements IEnumerable(Of TreeModel)

'    '表示名
'    Private _RealName As String
'    Public Property RealName As String
'        Get
'            Return Me._RealName
'        End Get
'        Set(value As String)
'            Me._RealName = value
'            RaisePropertyChanged("RealName")
'        End Set
'    End Property

'    '
'    Public ReadOnly Property Children As ObservableCollection(Of TreeModel) _
'        = New ObservableCollection(Of TreeModel)

'    '
'    Private _IsSelected As Boolean
'    Public Property IsSelected As Boolean
'        Get
'            Return Me._IsSelected
'        End Get
'        Set(value As Boolean)
'            Me._IsSelected = value
'            RaisePropertyChanged("_IsSelected")
'        End Set
'    End Property

'    '
'    Private _IsExpanded As Boolean
'    Public Property IsExpanded As Boolean
'        Get
'            Return Me._IsExpanded
'        End Get
'        Set(value As Boolean)
'            Me._IsExpanded = value
'            RaisePropertyChanged("_IsExpanded")
'        End Set
'    End Property

'    '
'    Private _IsChecked As Boolean
'    Public Property IsChecked As Boolean
'        Get
'            Return Me._IsChecked
'        End Get
'        Set(value As Boolean)
'            Me._IsChecked = value
'            RaisePropertyChanged("_IsChecked")
'        End Set
'    End Property

'    Friend Parent As TreeModel

'    '
'    Public Sub Add(c As TreeModel)
'        c.Parent = Me
'        Me.Children.Add(c)
'    End Sub

'    '
'    Public Sub Remove(c As TreeModel)
'        Me.Children.Remove(c)
'        c.Parent = Nothing
'    End Sub

'    '
'    Sub New(ByVal rn As String)
'        Me.RealName = rn
'    End Sub

'    Public Function GetEnumerator() As IEnumerator(Of TreeModel) Implements IEnumerable(Of TreeModel).GetEnumerator
'        Return Me.Children.GetEnumerator
'    End Function

'    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
'        Return Me.GetEnumerator()
'    End Function
'End Class
