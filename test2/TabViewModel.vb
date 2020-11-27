Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel

Public Class TabViewModel : Inherits BaseViewModel

    Private Property _Name As String
    Public Property Name As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
            RaisePropertyChanged("Name")
        End Set
    End Property

    Private Property _SelectedIndex As Integer
    Public Property SelectedIndex As Integer
        Get
            Return Me._SelectedIndex
        End Get
        Set(value As Integer)
            Me._SelectedIndex = value
            RaisePropertyChanged("SelectedIndex")
        End Set
    End Property

    ' 0...Left, 1...Top, 2...Right, 3...Bottom 
    Private _TabStripPlacement As String
    Public Property TabStripPlacement As String
        Get
            Return Me._TabStripPlacement
        End Get
        Set(value As String)
            Me._TabStripPlacement = value
            RaisePropertyChanged("TabStripPlacement")
        End Set
    End Property

    'Public ReadOnly Property RotateAngle As Double
    '    Get
    '        Select Case Me.TabStripPlacement
    '            Case "Bottom"
    '                Return 0.0
    '            Case "Left"
    '                Return 90.0
    '            Case "Right"
    '                Return 90.0
    '            Case "Top"
    '                Return 0.0
    '            Case Else
    '                Return 0.0
    '        End Select
    '    End Get
    'End Property


    Private _RotateAngleFirstCheck As Double
    Private _RotateAngle As Double
    Public Property RotateAngle As Double
        Get
            If Not Me._RotateAngleFirstCheck Then
                Select Case Me.TabStripPlacement
                    Case "Bottom"
                        Me._RotateAngle = 0.0
                    Case "Left"
                        Me._RotateAngle = 90.0
                    Case "Right"
                        Me._RotateAngle = 90.0
                    Case "Top"
                        Me._RotateAngle = 0.0
                    Case Else
                        Me._RotateAngle = 0.0
                End Select
                Me._RotateAngleFirstCheck = True
            End If
            Return Me._RotateAngle
        End Get
        Set(value As Double)
            Me._RotateAngle = value
        End Set
    End Property

    ' 廃止検討中
    'Private Property _SelectedItem As TabItemModel
    'Public Property SelectedItem As TabItemModel
    '    Get
    '        Return Me._SelectedItem
    '    End Get
    '    Set(value As TabItemModel)
    '        Me._SelectedItem = value
    '        RaisePropertyChanged("SelectedItem")
    '        Call _ChangeTabCloseButtonVisibilities(value)
    '        Call _ChangeTabColors(value)
    '    End Set
    'End Property

    Private Property _SelectedItem As ViewItemModel
    Public Property SelectedItem As ViewItemModel
        Get
            Return Me._SelectedItem
        End Get
        Set(value As ViewItemModel)
            Me._SelectedItem = value
            RaisePropertyChanged("SelectedItem")
            Call _ChangeCloseViewButtonVisibilities(value)
            'Call _ChangeTabColors(value)
        End Set
    End Property

    Private _ViewTabs As ObservableCollection(Of ViewItemModel)
    Public Property ViewTabs As ObservableCollection(Of ViewItemModel)
        Get
            If Me._ViewTabs Is Nothing Then
                Me._ViewTabs = New ObservableCollection(Of ViewItemModel)
            End If
            Return Me._ViewTabs
        End Get
        Set(value As ObservableCollection(Of ViewItemModel))
            Me._ViewTabs = value
        End Set
    End Property

    Private _PreservedViewTabs As ObservableCollection(Of ViewItemModel)
    Public Property PreservedViewTabs As ObservableCollection(Of ViewItemModel)
        Get
            If Me._PreservedViewTabs Is Nothing Then
                Me._PreservedViewTabs = New ObservableCollection(Of ViewItemModel)
            End If
            Return Me._PreservedViewTabs
        End Get
        Set(value As ObservableCollection(Of ViewItemModel))
            Me._PreservedViewTabs = value
        End Set
    End Property

    ' 廃止検討中
    'Private _Tabs As ObservableCollection(Of TabItemModel)
    'Public Property Tabs As ObservableCollection(Of TabItemModel)
    '    Get
    '        If Me._Tabs Is Nothing Then
    '            Me._Tabs = New ObservableCollection(Of TabItemModel)
    '        End If
    '        Return Me._Tabs
    '    End Get
    '    Set(value As ObservableCollection(Of TabItemModel))
    '        Me._Tabs = value
    '    End Set
    'End Property

    ' 廃止検討中
    'Private _PreservedTabs As ObservableCollection(Of TabItemModel)
    'Public Property PreservedTabs As ObservableCollection(Of TabItemModel)
    '    Get
    '        If Me._PreservedTabs Is Nothing Then
    '            Me._PreservedTabs = New ObservableCollection(Of TabItemModel)
    '        End If
    '        Return Me._PreservedTabs
    '    End Get
    '    Set(value As ObservableCollection(Of TabItemModel))
    '        Me._PreservedTabs = value
    '    End Set
    'End Property

    ' 廃止検討中
    'Private _ViewContentTabs As ObservableCollection(Of ViewItemModel)
    'Public Property ViewContentTabs As ObservableCollection(Of ViewItemModel)
    '    Get
    '        If Me._ViewContentTabs Is Nothing Then
    '            Me._ViewContentTabs = New ObservableCollection(Of ViewItemModel)
    '        End If
    '        Return Me._ViewContentTabs
    '    End Get
    '    Set(value As ObservableCollection(Of ViewItemModel))
    '        Me._ViewContentTabs = value
    '    End Set
    'End Property

    Private Sub _ChangeCloseViewButtonVisibilities(ByVal vim As ViewItemModel)
        If Me.ViewTabs IsNot Nothing Then
            For Each vt In Me.ViewTabs
                If vt.Equals(vim) Then
                    vt.CloseViewButtonVisibility = Visibility.Visible
                Else
                    vt.CloseViewButtonVisibility = Visibility.Collapsed
                End If
            Next
        End If
    End Sub

    Private Sub _ChangeTabColors(ByVal [tab] As TabItemModel)
        'If Me.Tabs IsNot Nothing Then
        '    For Each t In Me.Tabs
        '        If t.Equals([tab]) Then
        '            t.Color = Colors.Goldenrod
        '        Else
        '            t.Color = Colors.White
        '        End If
        '    Next
        'End If
    End Sub

    'Public Overloads Sub AddTab(ByVal [tab] As TabItemModel)
    '    Dim idx = -1
    '    For Each t In Me.Tabs
    '        If t.Name = [tab].Name Then
    '            idx = Me.Tabs.IndexOf(t)
    '            Exit For
    '        End If
    '    Next
    '    If idx > -1 Then
    '        Me.Tabs(idx) = [tab]
    '    Else
    '        Me.Tabs.Add([tab])
    '    End If

    '    For Each pt In Me.PreservedTabs
    '        If pt.Name = [tab].Name Then
    '            idx = Me.PreservedTabs.IndexOf(pt)
    '            Exit For
    '        End If
    '    Next
    '    If idx > -1 Then
    '        Me.PreservedTabs(idx) = [tab]
    '    Else
    '        Me.PreservedTabs.Add([tab])
    '    End If
    'End Sub

    'Public Overloads Sub AddTab(ByVal vim As ViewItemModel)
    '    Dim idx = -1

    '    vim.WrapperName = Me.Name

    '    For Each vt In Me.ViewContentTabs
    '        If vt.Name = vim.Name Then
    '            idx = Me.ViewContentTabs.IndexOf(vt)
    '            Exit For
    '        End If
    '    Next
    '    If idx > -1 Then
    '        Me.ViewContentTabs(idx) = vim
    '    Else
    '        Me.ViewContentTabs.Add(vim)
    '    End If

    '    Dim tim = New TabItemModel With {
    '        .Name = vim.Name,
    '        .ViewContent = vim
    '    }

    '    idx = -1
    '    For Each t In Me.Tabs
    '        If t.Name = vim.Name Then
    '            idx = Me.Tabs.IndexOf(t)
    '            Exit For
    '        End If
    '    Next
    '    If idx > -1 Then
    '        Me.Tabs(idx) = tim
    '    Else
    '        Me.Tabs.Add(tim)
    '    End If
    'End Sub

    Public Overloads Sub AddTab(ByVal vim As ViewItemModel)
        Dim idx = -1
        For Each vt In Me.ViewTabs
            If vt.Name = vim.Name Then
                idx = Me.ViewTabs.IndexOf(vt)
                Exit For
            End If
        Next
        If idx > -1 Then
            Me.ViewTabs(idx) = vim
        Else
            Me.ViewTabs.Add(vim)
        End If

        idx = -1
        For Each pvt In Me.PreservedViewTabs
            If pvt.Name = vim.Name Then
                idx = Me.PreservedViewTabs.IndexOf(pvt)
                Exit For
            End If
        Next
        If idx > -1 Then
            Me.PreservedViewTabs(idx) = vim
        Else
            Me.PreservedViewTabs.Add(vim)
        End If
    End Sub

    '--- タブを閉じる関連 ------------------------------------------------------------------------'
    Private Sub _TabCloseRequestedReview(ByVal t As TabItemModel, ByVal e As System.EventArgs)
        'If Me.Tabs.Contains(t) Then
        '    Call _TabCloseRequestAccept(t)
        'End If
    End Sub

    Private Sub _TabCloseRequestAccept(ByVal [tab] As TabItemModel)
        'Me.ViewTabs.Remove([tab])

        'For Each vt In Me.ViewContentTabs
        '    If vt.Name = [tab].ViewContent.Name Then
        '        vt.IsVisible = False
        '        Exit For
        '    End If
        'Next

        'If Me.Tabs.Count = 0 Then
        '    Call DelegateEventListener.Instance.RaiseTabCloseViewd()
        'End If
    End Sub

    Private Sub _TabCloseAddHandler()
        AddHandler _
            DelegateEventListener.Instance.TabCloseRequested,
            AddressOf Me._TabCloseRequestedReview
    End Sub
    '---------------------------------------------------------------------------------------------'

    Public Sub New()
        Call _TabCloseAddHandler()
    End Sub
End Class
