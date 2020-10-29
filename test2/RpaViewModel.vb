Imports test2

Public Class RpaViewModel : Inherits BaseViewModel2
    Private __Rpa As RpaModel
    Private Property _Rpa As RpaModel
        Get
            Return Me.__Rpa
        End Get
        Set(value As RpaModel)
            Me.__Rpa = value
            AppInfo.ProjectInfo.Model.Data.Rpas(
                AppInfo.ProjectInfo.Model.Data.Rpas.IndexOf(Me._Rpa)
            ) = value
        End Set
    End Property

    Private _RootProjectName As String
    Public Property RootProjectName As String
        Get
            Return Me._RootProjectName
        End Get
        Set(value As String)
            Me._RootProjectName = value
            Me._Rpa.RootProjectName = value
            RaisePropertyChanged("RootProjectName")
        End Set
    End Property

    Private _UserProjectName As String
    Public Property UserProjectName As String
        Get
            Return Me._UserProjectName
        End Get
        Set(value As String)
            Me._UserProjectName = value
            Me._Rpa.UserProjectName = value
            RaisePropertyChanged("UserProjectName")
        End Set
    End Property

    Private _Status As Integer
    Public Property Status As Integer
        Get
            Return Me._Status
        End Get
        Set(value As Integer)
            Me._Status = value
            Me._Rpa.Status = value
            RaisePropertyChanged("StatusText")
        End Set
    End Property

    Private _StatusText As String
    Public ReadOnly Property StatusText As String
        Get
            Select Case Me.Status
                Case 0
                    Me._StatusText = "未接続"
                Case 1
                    Me._StatusText = "接続済み"
            End Select
            Return Me._StatusText
        End Get
    End Property

    Private _RunParameter As String
    Public Property RunParameter As String
        Get
            Return Me._RunParameter
        End Get
        Set(value As String)
            Me._RunParameter = value
            Me._Rpa.RunParameter = value
            RaisePropertyChanged("RunParameter")
        End Set
    End Property

    Public Sub _ViewInitializing()
        Dim b = False
        Dim rpa As RpaModel
        For Each r In AppInfo.ProjectInfo.Model.Data.Rpas
            If Not r.IsViewAssigned Then
                Me._Rpa = r
                Me.RootProjectName = r.RootProjectName
                Me.UserProjectName = r.UserProjectName
                Me.Status = r.Status
                Me.RunParameter = r.RunParameter
                r.IsViewAssigned = True
                b = True
                Exit For
            End If
        Next
        If Not b Then
            Me._Rpa = New RpaModel With {
                .Index = (AppInfo.ProjectInfo.Model.Data.GetRpaIndex()),
                .IsViewAssigned = True
            }
            Me.RootProjectName = vbNullString
            Me.UserProjectName = vbNullString
            AppInfo.ProjectInfo.Model.Data.Rpas.Add(Me._Rpa)
        End If
    End Sub

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        InitializeHandler = AddressOf _ViewInitializing
        Call BaseInitialize(app, vm)
    End Sub
End Class
