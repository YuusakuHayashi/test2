﻿Imports System.Collections.ObjectModel

Public Class ProjectExplorerViewModel : Inherits ControllerViewModelBase(Of ProjectExplorerViewModel)
    Private ReadOnly Property ProjectExplorerProjectIconFileName As String
        Get
            Return $"{Controller.SystemDllDirectory}\projectexplorerproject.png"
        End Get
    End Property
    Private _ProjectExplorerProjectIcon As BitmapImage
    Public ReadOnly Property ProjectExplorerProjectIcon As BitmapImage
        Get
            If Me._ProjectExplorerProjectIcon Is Nothing Then
                Me._ProjectExplorerProjectIcon = RpaGuiModule.CreateIcon(Me.ProjectExplorerProjectIconFileName)
            End If
            Return Me._ProjectExplorerProjectIcon
        End Get
    End Property

    Private ReadOnly Property ProjectExplorerRobotIconFileName As String
        Get
            Return $"{Controller.SystemDllDirectory}\projectexplorerrobot.png"
        End Get
    End Property
    Private _ProjectExplorerRobotIcon As BitmapImage
    Public ReadOnly Property ProjectExplorerRobotIcon As BitmapImage
        Get
            If Me._ProjectExplorerRobotIcon Is Nothing Then
                Me._ProjectExplorerRobotIcon = RpaGuiModule.CreateIcon(Me.ProjectExplorerRobotIconFileName)
            End If
            Return Me._ProjectExplorerRobotIcon
        End Get
    End Property

    Public Class ProjectData : Inherits ViewModelBase
        Private _Icon As BitmapImage
        Public Property Icon As BitmapImage
            Get
                Return Me._Icon
            End Get
            Set(value As BitmapImage)
                Me._Icon = value
                RaisePropertyChanged("Icon")
            End Set
        End Property

        Private _Name As String
        Public Property Name As String
            Get
                Return Me._Name
            End Get
            Set(value As String)
                Me._Name = value
                RaisePropertyChanged("Name")
            End Set
        End Property

        Private _Status As Integer
        Public Property Status As Integer
            Get
                Return Me._Status
            End Get
            Set(value As Integer)
                Me._Status = value
                RaisePropertyChanged("Status")
            End Set
        End Property

        Private _Robots As ObservableCollection(Of RobotData)
        Public Property Robots As ObservableCollection(Of RobotData)
            Get
                If Me._Robots Is Nothing Then
                    Me._Robots = New ObservableCollection(Of RobotData)
                End If
                Return Me._Robots
            End Get
            Set(value As ObservableCollection(Of RobotData))
                Me._Robots = value
                RaisePropertyChanged("Robots")
            End Set
        End Property

        Public Class RobotData : Inherits ViewModelBase
            Private _Icon As BitmapImage
            Public Property Icon As BitmapImage
                Get
                    Return Me._Icon
                End Get
                Set(value As BitmapImage)
                    Me._Icon = value
                    RaisePropertyChanged("Icon")
                End Set
            End Property

            Private _Name As String
            Public Property Name As String
                Get
                    Return Me._Name
                End Get
                Set(value As String)
                    Me._Name = value
                    RaisePropertyChanged("Name")
                End Set
            End Property

            Private _Status As Integer
            Public Property Status As Integer
                Get
                    Return Me._Status
                End Get
                Set(value As Integer)
                    Me._Status = value
                    RaisePropertyChanged("Status")
                End Set
            End Property
        End Class
    End Class

    Private _Projects As ObservableCollection(Of ProjectData)
    Public Property Projects As ObservableCollection(Of ProjectData)
        Get
            If Me._Projects Is Nothing Then
                Me._Projects = New ObservableCollection(Of ProjectData)
            End If
            Return Me._Projects
        End Get
        Set(value As ObservableCollection(Of ProjectData))
            Me._Projects = value
            RaisePropertyChanged("Projects")
        End Set
    End Property

    Public Overrides Sub Initialize()
        For Each project In ViewController.Data.Initializer.Projects
            Dim pd As New ProjectData
            pd.Icon = Me.ProjectExplorerProjectIcon
            pd.Name = project.Name
            Me.Projects.Add(pd)
        Next
    End Sub
End Class
