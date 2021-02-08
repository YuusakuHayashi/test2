Public Class RunnerViewModel : Inherits ExecutableMenuBase
    Private _RunRobotIcon As BitmapImage
    Public Property [RunRobotIcon] As BitmapImage
        Get
            If Me._RunRobotIcon Is Nothing Then
                Me._RunRobotIcon = CreateIcon($"{Controller.SystemDllDirectory}\runrobot.png")
            End If
            Return _RunRobotIcon
        End Get
        Set(value As BitmapImage)
            Me._RunRobotIcon = value
            RaisePropertyChanged("RunRobotIcon")
        End Set
    End Property

    Private _ProjectName As String
    Public Property ProjectName As String
        Get
            Return Me._ProjectName
        End Get
        Set(value As String)
            Me._ProjectName = value
            RaisePropertyChanged("ProjectName")
        End Set
    End Property

    Private _RobotName As String
    Public Property RobotName As String
        Get
            Return Me._RobotName
        End Get
        Set(value As String)
            Me._RobotName = value
            RaisePropertyChanged("RobotName")
        End Set
    End Property

    Private _RobotAlias As String
    Public Property RobotAlias As String
        Get
            Return Me._RobotAlias
        End Get
        Set(value As String)
            Me._RobotAlias = value
            RaisePropertyChanged("RobotAlias")
        End Set
    End Property

    Private _RootRobotDirectory As String
    Public Property RootRobotDirectory As String
        Get
            Return Me._RootRobotDirectory
        End Get
        Set(value As String)
            Me._RootRobotDirectory = value
            RaisePropertyChanged("RootRobotDirectory")
        End Set
    End Property

    Private _MyRobotDirectory As String
    Public Property MyRobotDirectory As String
        Get
            Return Me._MyRobotDirectory
        End Get
        Set(value As String)
            Me._MyRobotDirectory = value
            RaisePropertyChanged("MyRobotDirectory")
        End Set
    End Property

    Private _MyRobotReadMe As String
    Public Property MyRobotReadMe As String
        Get
            Return Me._MyRobotReadMe
        End Get
        Set(value As String)
            Me._MyRobotReadMe = value
            RaisePropertyChanged("MyRobotReadMe")
        End Set
    End Property

    '---------------------------------------------------------------------------------------------'
    Private _RunRobotCommand As ICommand
    Public ReadOnly Property RunRobotCommand As ICommand
        Get
            If Me._RunRobotCommand Is Nothing Then
                Me._RunRobotCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _RunRobotCommandExecute,
                    .CanExecuteHandler = AddressOf _RunRobotCommandCanExecute
                }
                Return Me._RunRobotCommand
            Else
                Return Me._RunRobotCommand
            End If
        End Get
    End Property

    Private __RunRobotCommandEnableFlag As Boolean
    Public Property _RunRobotCommandEnableFlag As Boolean
        Get
            Return Me.__RunRobotCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__RunRobotCommandEnableFlag = value
            RaisePropertyChanged("_RunRobotCommandEnableFlag")
            CType(RunRobotCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Async Sub _RunRobotCommandExecute(ByVal parameter As Object)
        Me.Data.System.ExecuteMode = Rpa00.RpaSystem.ExecuteModeNumber.RpaGui
        Me.Data.System.GuiCommandText = "runrobot noprint end"
        Dim t As Task(Of String) = Task.Run(
            Function()
                Data.System.Output.OutputText &= "hoge"
                Return "hoge"
            End Function
        )
        Await t
        Threading.Thread.Sleep(1000)
        Data.System.Output.OutputText &= t.Result

        Threading.Thread.Sleep(1000)
        Dim t2 As Task(Of String) = Task.Run(
            Function()
                Data.System.Output.OutputText &= "fuga"
                Return "fuga"
            End Function
        )
        Await t2
        Data.System.Output.OutputText &= t2.Result
        Call Data.System.CuiLoop(Data)
    End Sub

    Private Function _RunRobotCommandCanExecute(ByVal parameter As Object) As Boolean
        Return True
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Function CreateIcon(ByVal filename As String) As BitmapImage
        Dim bi = New BitmapImage
        bi.BeginInit()
        bi.UriSource = New Uri(filename, UriKind.Absolute)
        bi.EndInit()
        Return bi
    End Function

    Sub Test()
        '重い処理
    End Sub

    Sub New(ByRef dat As Object)
        Me.Data = dat
        If Data IsNot Nothing Then
            If Data.Project IsNot Nothing Then
                Me.ProjectName = Data.Project.ProjectName
                Me.RobotName = Data.Project.RobotName
                Me.RobotAlias = Data.Project.RobotAlias
                Me.MyRobotDirectory = Data.Project.MyRobotDirectory
                Me.RootRobotDirectory = Data.Project.RootRobotDirectory
                Me.MyRobotReadMe = Data.Project.MyRobotReadMe
            End If
        End If
    End Sub
End Class
