Imports System.ComponentModel
Imports System.IO
Imports System.Collections.ObjectModel

Public Class DirectoryInfoModel
    Implements INotifyPropertyChanged

    ' INortify
    '-------------------------------------------------------------------------'
    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub
    '-------------------------------------------------------------------------'
    Private _Icon As BitmapImage
    Public Property Icon As BitmapImage
        Get
            Return _Icon
        End Get
        Set(value As BitmapImage)
            _Icon = value
            RaisePropertyChanged("Icon")
        End Set
    End Property

    Public ReadOnly Property FileType As String
        Get
            If File.Exists(Me.FullName) Then
                Return "File"
            ElseIf Directory.Exists(Me.FullName) Then
                Return "Folder"
            Else
                Return "Unknown"
            End If
        End Get
    End Property

    Public ReadOnly Property Name As String
        Get
            Return Path.GetFileName(Me.FullName)
        End Get
    End Property

    Private _FullName As String
    Public Property FullName As String
        Get
            Return _FullName
        End Get
        Set(value As String)
            _FullName = value
            Call _CreateTransparentChildren()
            RaisePropertyChanged("FullName")
        End Set
    End Property

    Private _IsSelected As Boolean
    Public Property IsSelected As Boolean
        Get
            Return _IsSelected
        End Get
        Set(value As Boolean)
            _IsSelected = value
            RaisePropertyChanged("IsSelected")
        End Set
    End Property

    Private _IsExpanded As Boolean
    Public Property IsExpanded As Boolean
        Get
            Return Me._IsExpanded
        End Get
        Set(value As Boolean)
            _IsExpanded = value
            If value Then
                Call _CreateChildren()
            Else
                Call _CreateTransparentChildren()
            End If
            RaisePropertyChanged("IsExpanded")
        End Set
    End Property

    Private _Children As ObservableCollection(Of DirectoryInfoModel)
    Public Property Children As ObservableCollection(Of DirectoryInfoModel)
        Get
            Return _Children
        End Get
        Set(value As ObservableCollection(Of DirectoryInfoModel))
            _Children = value
            RaisePropertyChanged("Children")
        End Set
    End Property

    ' 実際に展開した際に、空の子を殺して、本当の子データを作成する。
    ' 展開するのはディレクトリの時だけなので、タイプは意識しなくてよい
    Private Sub _CreateChildren()
        Me.Children = New ObservableCollection(Of DirectoryInfoModel)
        For Each child In Directory.GetFileSystemEntries(Me.FullName)
            Me.Children.Add(
                New DirectoryInfoModel With {
                    .FullName = child
                }
            )
        Next
        If Me.Children.Count = 0 Then
            Me.Children.Add(New DirectoryInfoModel With {.FullName = "(ファイルなし)"})
        End If
    End Sub

    ' エクスパンダーを表示するために、空の子を作成する。
    Private Sub _CreateTransparentChildren()
        Me.Children = New ObservableCollection(Of DirectoryInfoModel)
        If Me.FileType = "Folder" Then
            Me.Children.Add(New DirectoryInfoModel)
        End If
    End Sub
End Class
