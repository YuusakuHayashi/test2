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

    Private Property _SelectedItem As TabItemModel
    Public Property SelectedItem As TabItemModel
        Get
            Return Me._SelectedItem
        End Get
        Set(value As TabItemModel)
            Me._SelectedItem = value
            RaisePropertyChanged("SelectedItem")
            Call _ChangeTabCloseButtonVisibilities(value)
            Call _ChangeTabColors(value)
        End Set
    End Property

    Private _Tabs As ObservableCollection(Of TabItemModel)
    Public Property Tabs As ObservableCollection(Of TabItemModel)
        Get
            If Me._Tabs Is Nothing Then
                Me._Tabs = New ObservableCollection(Of TabItemModel)
            End If
            Return Me._Tabs
        End Get
        Set(value As ObservableCollection(Of TabItemModel))
            Me._Tabs = value
        End Set
    End Property

    Private _PreservedTabs As ObservableCollection(Of TabItemModel)
    Public Property PreservedTabs As ObservableCollection(Of TabItemModel)
        Get
            If Me._PreservedTabs Is Nothing Then
                Me._PreservedTabs = New ObservableCollection(Of TabItemModel)
            End If
            Return Me._PreservedTabs
        End Get
        Set(value As ObservableCollection(Of TabItemModel))
            Me._PreservedTabs = value
        End Set
    End Property

    Private _ViewContentTabs As ObservableCollection(Of ViewItemModel)
    Public Property ViewContentTabs As ObservableCollection(Of ViewItemModel)
        Get
            If Me._ViewContentTabs Is Nothing Then
                Me._ViewContentTabs = New ObservableCollection(Of ViewItemModel)
            End If
            Return Me._ViewContentTabs
        End Get
        Set(value As ObservableCollection(Of ViewItemModel))
            Me._ViewContentTabs = value
        End Set
    End Property

    Private Sub _ChangeTabCloseButtonVisibilities(ByVal [tab] As TabItemModel)
        If Me.Tabs IsNot Nothing Then
            For Each t In Me.Tabs
                If t.Equals([tab]) Then
                    t.TabCloseButtonVisibility = Visibility.Visible
                Else
                    t.TabCloseButtonVisibility = Visibility.Collapsed
                End If
            Next
        End If
    End Sub

    Private Sub _ChangeTabColors(ByVal [tab] As TabItemModel)
        If Me.Tabs IsNot Nothing Then
            For Each t In Me.Tabs
                If t.Equals([tab]) Then
                    t.Color = Colors.Goldenrod
                Else
                    t.Color = Colors.White
                End If
            Next
        End If
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

    Public Overloads Sub AddTab(ByVal vim As ViewItemModel)
        Dim idx = -1

        vim.WrapperName = Me.Name

        For Each vt In Me.ViewContentTabs
            If vt.Name = vim.Name Then
                idx = Me.ViewContentTabs.IndexOf(vt)
                Exit For
            End If
        Next
        If idx > -1 Then
            Me.ViewContentTabs(idx) = vim
        Else
            Me.ViewContentTabs.Add(vim)
        End If

        Dim tim = New TabItemModel With {
            .Name = vim.Name,
            .ViewContent = vim
        }

        idx = -1
        For Each t In Me.Tabs
            If t.Name = vim.Name Then
                idx = Me.Tabs.IndexOf(t)
                Exit For
            End If
        Next
        If idx > -1 Then
            Me.Tabs(idx) = tim
        Else
            Me.Tabs.Add(tim)
        End If
    End Sub

    '--- タブを閉じる関連 ------------------------------------------------------------------------'
    Private Sub _TabCloseRequestedReview(ByVal t As TabItemModel, ByVal e As System.EventArgs)
        If Me.Tabs.Contains(t) Then
            Call _TabCloseRequestAccept(t)
        End If
    End Sub

    Private Sub _TabCloseRequestAccept(ByVal [tab] As TabItemModel)
        Me.Tabs.Remove([tab])
        For Each vt In Me.ViewContentTabs
            If vt.Name = [tab].ViewContent.Name Then
                vt.IsVisible = False
                Exit For
            End If
        Next

        If Me.Tabs.Count = 0 Then
            Call DelegateEventListener.Instance.RaiseTabViewClosed()
        End If
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
