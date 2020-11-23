Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Collections.ObjectModel

Public Class ViewItemModel
    Inherits BaseViewModel

    Private _IsVisible As Boolean
    Public Property IsVisible As Boolean
        Get
            Return _IsVisible
        End Get
        Set(value As Boolean)
            _IsVisible = value
        End Set
    End Property

    Private _VisibleIcon As BitmapImage
    <JsonIgnore>
    Public Property VisibleIcon As BitmapImage
        Get
            Return _VisibleIcon
        End Get
        Set(value As BitmapImage)
            _VisibleIcon = value
            RaisePropertyChanged("VisibleIcon")
        End Set
    End Property

    Private _IconFileName As String
    Public Property IconFileName As String
        Get
            Return _IconFileName
        End Get
        Set(value As String)
            _IconFileName = value
        End Set
    End Property

    Private _Icon As BitmapImage
    <JsonIgnore>
    Public Property [Icon] As BitmapImage
        Get
            Return _Icon
        End Get
        Set(value As BitmapImage)
            _Icon = value
        End Set
    End Property

    Private _BoxVisibility As Visibility
    <JsonIgnore>
    Public Property BoxVisibility As Visibility
        Get
            Return Me._BoxVisibility
        End Get
        Set(value As Visibility)
            Me._BoxVisibility = value
            If value = Visibility.Visible Then
                Me.BlockVisibility = Visibility.Collapsed
            End If
            RaisePropertyChanged("BlockVisibility")
        End Set
    End Property

    Private _BlockVisibility As Visibility
    <JsonIgnore>
    Public Property BlockVisibility As Visibility
        Get
            Return Me._BlockVisibility
        End Get
        Set(value As Visibility)
            Me._BlockVisibility = value
            If value = Visibility.Visible Then
                Me.BoxVisibility = Visibility.Collapsed
            End If
            RaisePropertyChanged("BoxVisibility")
        End Set
    End Property

    Private _Name As String
    Public Property Name As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
        End Set
    End Property

    Private _Alias As String
    Public Property [Alias] As String
        Get
            If String.IsNullOrEmpty(Me._Alias) Then
                Me._Alias = Me.Name
            End If
            Return Me._Alias
        End Get
        Set(value As String)
            Me._Alias = value
            If BoxVisibility = Visibility.Visible Then
                BlockVisibility = Visibility.Visible
            End If
            RaisePropertyChanged("Alias")
        End Set
    End Property

    Private _ModelName As String
    Public Property ModelName As String
        Get
            Return Me._ModelName
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(Me._ModelName) Then
                Me._ModelName = value
            End If
        End Set
    End Property

    Private _FrameName As String
    Public Property FrameName As String
        Get
            Return Me._FrameName
        End Get
        Set(value As String)
            Me._FrameName = value
        End Set
    End Property

    Private _OpenState As Boolean
    Public Property OpenState As Boolean
        Get
            Return Me._OpenState
        End Get
        Set(value As Boolean)
            Me._OpenState = value
        End Set
    End Property

    Private _Content As Object
    Public Property Content As Object
        Get
            Return Me._Content
        End Get
        Set(value As Object)
            Me._Content = value
            If value Is Nothing Then
                Me.ModelName = Nothing
            Else
                Me.ModelName = value.GetType.Name
            End If
            RaisePropertyChanged("Content")
        End Set
    End Property

    Private _CloseViewButtonVisibility As Visibility
    <JsonIgnore>
    Public Property CloseViewButtonVisibility As Visibility
        Get
            Return Me._CloseViewButtonVisibility
        End Get
        Set(value As Visibility)
            Me._CloseViewButtonVisibility = value
            RaisePropertyChanged("CloseViewButtonVisibility")
        End Set
    End Property

    ' タブ関連
    '---------------------------------------------------------------------------------------------'
    Private _Parent As ViewItemModel
    <JsonIgnore>
    Public Property Parent As ViewItemModel
        Get
            Return Me._Parent
        End Get
        Set(value As ViewItemModel)
            Me._Parent = value
        End Set
    End Property

    Private _Children As ObservableCollection(Of ViewItemModel)
    Public Property Children As ObservableCollection(Of ViewItemModel)
        Get
            If Me._Children Is Nothing Then
                Me._Children = New ObservableCollection(Of ViewItemModel)
            End If
            Return Me._Children
        End Get
        Set(value As ObservableCollection(Of ViewItemModel))
            Me._Children = value
        End Set
    End Property

    ' ContentがTabViewModelの時に利用可能
    Public Sub RegisterChildren()
        Dim idx = -1
        If Me.ModelName = "TabViewModel" Then
            For Each child In Me.Content.ViewTabs
                ' 親の認知
                child.Parent = Me
                Me.Children.Add(child)
            Next
        End If
    End Sub
    '---------------------------------------------------------------------------------------------'

    'Private _IsExpand As Boolean
    'Public Property IsExpand As Boolean
    '    Get
    '        Return Me._IsExpand
    '    End Get
    '    Set(value As Boolean)
    '        Me._IsExpand = value
    '    End Set
    'End Property

    Private _IsSelected As Boolean
    <JsonIgnore>
    Public Property IsSelected As Boolean
        Get
            Return Me._IsSelected
        End Get
        Set(value As Boolean)
            Me._IsSelected = value
            'If value Then
            '    'Call DelegateEventListener.Instance.RaiseOpenViewRequested(Me)
            'End If
        End Set
    End Property

    'Private _IsReadOnly As Boolean
    'Public Property IsReadOnly As Boolean
    '    Get
    '        Return Me._IsReadOnly
    '    End Get
    '    Set(value As Boolean)
    '        Me._IsReadOnly = value
    '        RaisePropertyChanged("IsReadOnly")
    '    End Set
    'End Property

    ' コマンドプロパティ（タブ閉じる）
    '---------------------------------------------------------------------------------------------'
    Private _CloseViewCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property CloseViewCommand As ICommand
        Get
            If Me._CloseViewCommand Is Nothing Then
                Me._CloseViewCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _CloseViewCommandExecute,
                    .CanExecuteHandler = AddressOf _CloseViewCommandCanExecute
                }
                Return Me._CloseViewCommand
            Else
                Return Me._CloseViewCommand
            End If
        End Get
    End Property

    'コマンド実行可否のチェック（タブ閉じる）
    Private Sub _CheckCloseViewCommandEnabled()
        Dim b = True
        Me._CloseViewCommandEnableFlag = b
    End Sub

    ' コマンド実行可否のフラグ（タブ閉じる）
    Private __CloseViewCommandEnableFlag As Boolean
    Private Property _CloseViewCommandEnableFlag As Boolean
        Get
            Return Me.__CloseViewCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__CloseViewCommandEnableFlag = value
            RaisePropertyChanged("_CloseViewCommandEnableFlag")
            CType(CloseViewCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    ' コマンド実行（タブ閉じる）
    Private Sub _CloseViewCommandExecute(ByVal parameter As Object)
        Me.IsVisible = False
        Me.Parent.Content.ViewTabs.Remove(Me)
        If Me.Parent.Content.ViewTabs.Count = 0 Then
            Me.Parent.Content = Nothing
            Me.Parent.IsVisible = False
            Call DelegateEventListener.Instance.RaiseCloseViewRequested(Me.Parent)
        End If
        Call DelegateEventListener.Instance.RaiseReloadViewsRequested()
    End Sub

    Private Function _CloseViewCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._CloseViewCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    ' 閉じるコマンド受理
    '---------------------------------------------------------------------------------------------'
    'Private Sub _CloseViewRequestedAddHandler()
    '    AddHandler _
    '        DelegateEventListener.Instance.CloseViewRequested,
    '        AddressOf Me._CloseViewRequestedReview
    'End Sub

    'Private Sub _CloseViewRequestedReview(ByVal child As ViewItemModel, System.EventArgs)
    '    If child.Parent.Equals(Me) Then
    '        Call _CloseViewRequestedAccept(child)
    '    End If
    'End Sub

    'Private Sub _CloseViewRequestedAccept(ByVal sender As Object)
    '    Select Case sender.Name
    '        Case "ContentPanel"
    '            sender.DataContext.ContentViewWidth = sender.ActualWidth
    '            sender.DataContext.ContentViewHeight = sender.ActualHeight
    '        Case "RightPanel"
    '            sender.DataContext.RightViewWidth = sender.ActualWidth
    '        Case "BottomPanel"
    '            sender.DataContext.BottomViewHeight = sender.ActualHeight
    '    End Select
    'End Sub
    '---------------------------------------------------------------------------------------------'


    '---------------------------------------------------------------------------------------------'
    ' コマンドプロパティ（接続確認）
    Private _ChangeAliasCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property ChangeAliasCommand As ICommand
        Get
            If Me._ChangeAliasCommand Is Nothing Then
                Me._ChangeAliasCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ChangeAliasCommandExecute,
                    .CanExecuteHandler = AddressOf _ChangeAliasCommandCanExecute
                }
                Return Me._ChangeAliasCommand
            Else
                Return Me._ChangeAliasCommand
            End If
        End Get
    End Property

    'コマンド実行可否のチェック（接続確認）
    Private Sub _CheckChangeAliasCommandEnabled()
        Dim b As Boolean : b = True
        Me._ChangeAliasCommandEnableFlag = b
    End Sub

    ' コマンド実行可否のフラグ（接続確認）
    Private __ChangeAliasCommandEnableFlag As Boolean
    <JsonIgnore>
    Public Property _ChangeAliasCommandEnableFlag As Boolean
        Get
            Return Me.__ChangeAliasCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ChangeAliasCommandEnableFlag = value
            RaisePropertyChanged("_ChangeAliasCommandEnableFlag")
            CType(ChangeAliasCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    ' コマンド実行（接続確認）
    Private Sub _ChangeAliasCommandExecute(ByVal parameter As Object)
        Me.BoxVisibility = Visibility.Visible
    End Sub

    Private Function _ChangeAliasCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ChangeAliasCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    ' コマンドプロパティ（ビュー開く）
    '---------------------------------------------------------------------------------------------'
    Private _OpenViewCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property OpenViewCommand As ICommand
        Get
            If Me._OpenViewCommand Is Nothing Then
                Me._OpenViewCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _OpenViewCommandExecute,
                    .CanExecuteHandler = AddressOf _OpenViewCommandCanExecute
                }
                Return Me._OpenViewCommand
            Else
                Return Me._OpenViewCommand
            End If
        End Get
    End Property

    'コマンド実行可否のチェック（ビュー削除）
    Private Sub _CheckOpenViewCommandEnabled()
        Dim b = True
        Me._OpenViewCommandEnableFlag = b
    End Sub

    ' コマンド実行可否のフラグ（ビュー削除）
    Private __OpenViewCommandEnableFlag As Boolean
    <JsonIgnore>
    Public Property _OpenViewCommandEnableFlag As Boolean
        Get
            Return Me.__OpenViewCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__OpenViewCommandEnableFlag = value
            RaisePropertyChanged("_OpenViewCommandEnableFlag")
            CType(OpenViewCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    ' コマンド実行（ビュー削除）
    Private Sub _OpenViewCommandExecute(ByVal parameter As Object)
        Call DelegateEventListener.Instance.RaiseOpenViewRequested(Me)
        Call DelegateEventListener.Instance.RaiseReloadViewsRequested()
    End Sub

    Private Function _OpenViewCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._OpenViewCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    Public Sub New()
        Call _CheckOpenViewCommandEnabled()
        Call _CheckCloseViewCommandEnabled()
        Call _CheckChangeAliasCommandEnabled()
        Me.BlockVisibility = Visibility.Visible
        Me.IsVisible = True
    End Sub
End Class
