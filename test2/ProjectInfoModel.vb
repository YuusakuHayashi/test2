Imports System.Windows.Forms
Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.ComponentModel

Public Class ProjectInfoModel
    Inherits JsonHandler(Of ProjectInfoModel)
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

    Private Const SHIFT_JIS As String = "Shift-JIS"

    Private _Model As Model
    <JsonIgnore>
    Public Property Model As Model
        Get
            If Me._Model Is Nothing Then
                Me._Model = New Model
            End If
            Return _Model
        End Get
        Set(value As Model)
            _Model = value
        End Set
    End Property

    Private _DirectoryInfo As DirectoryInfoModel
    <JsonIgnore>
    Public Property DirectoryInfo As DirectoryInfoModel
        Get
            If Me._DirectoryInfo Is Nothing Then
                If Not String.IsNullOrEmpty(Me.DirectoryName) Then
                    Me._DirectoryInfo = New DirectoryInfoModel With {
                        .Name = Me.DirectoryName
                    }
                End If
            End If
            Return Me._DirectoryInfo
        End Get
        Set(value As DirectoryInfoModel)
            Me._DirectoryInfo = value
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

    Private _Index As Integer
    Public Property [Index] As Integer
        Get
            Return _Index
        End Get
        Set(value As Integer)
            _Index = value
        End Set
    End Property

    Private _FixedIndex As Integer
    Public Property [FixedIndex] As Integer
        Get
            Return _FixedIndex
        End Get
        Set(value As Integer)
            _FixedIndex = value
        End Set
    End Property

    Private _Name As String
    Public Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            _Name = value
        End Set
    End Property

    Private _Kind As String
    Public Property Kind As String
        Get
            Return _Kind
        End Get
        Set(value As String)
            _Kind = value
        End Set
    End Property

    Private _LastUpdate As String
    Public Property LastUpdate As String
        Get
            Return _LastUpdate
        End Get
        Set(value As String)
            _LastUpdate = value
        End Set
    End Property

    Private _RootDirectoryName As String
    Public Property RootDirectoryName As String
        Get
            Return _RootDirectoryName
        End Get
        Set(value As String)
            _RootDirectoryName = value
        End Set
    End Property

    Private _DirectoryName As String
    Public Property DirectoryName As String
        Get
            Return _DirectoryName
        End Get
        Set(value As String)
            _DirectoryName = value
        End Set
    End Property

    Private _IsSelected As Boolean
    <JsonIgnore>
    Public Property IsSelected As Boolean
        Get
            Return Me._IsSelected
        End Get
        Set(value As Boolean)
            Me._IsSelected = value
            If value Then
                Call DelegateEventListener.Instance.RaiseOpenProjectRequested(Me)
            End If
        End Set
    End Property

    Private _IsExpand As Boolean
    <JsonIgnore>
    Public Property IsExpand As Boolean
        Get
            Return _IsExpand
        End Get
        Set(value As Boolean)
            _IsExpand = value
        End Set
    End Property

    Private _ActiveStatus As String
    <JsonIgnore>
    Public Property ActiveStatus As String
        Get
            Return _ActiveStatus
        End Get
        Set(value As String)
            _ActiveStatus = value
            RaisePropertyChanged("ActiveStatus")
        End Set
    End Property

    ' プロジェクトのモデルファイルです
    Private _ModelFileName As String
    Public Property ModelFileName As String
        Get
            If String.IsNullOrEmpty(Me._ModelFileName) Then
                Me._ModelFileName = Me.DirectoryName & "\Model.json"
            End If
            Return Me._ModelFileName
        End Get
        Set(value As String)
            Me._ModelFileName = value
            RaisePropertyChanged("ModelFileName")
        End Set
    End Property

    ' プロジェクトがアプリケーションにより作成されたことを表すファイルです
    Public ReadOnly Property IniFileName As String
        Get
            Return DirectoryName & "\Project.ini"
        End Get
    End Property

    Private _ProjectInfoFileName As String
    Public Property ProjectInfoFileName As String
        Get
            If String.IsNullOrEmpty(Me._ProjectInfoFileName) Then
                Me._ProjectInfoFileName = Me.DirectoryName & "\ProjectInfo.json"
            End If
            Return Me._ProjectInfoFileName
        End Get
        Set(value As String)
            Me._ProjectInfoFileName = value
            RaisePropertyChanged("ProjectInfoFileName")
        End Set
    End Property

    Private _ViewModelFileName As String
    Public Property ViewModelFileName As String
        Get
            If String.IsNullOrEmpty(Me._ViewModelFileName) Then
                Me._ViewModelFileName = Me.DirectoryName & "\ViewModel.json"
            End If
            Return Me._ViewModelFileName
        End Get
        Set(value As String)
            Me._ViewModelFileName = value
            RaisePropertyChanged("ViewModelFileName")
        End Set
    End Property

    'Public Sub ProjectSave()
    '    Me.ModelSave(Me.ProjectInfoFileName, Me)
    'End Sub

    Private Sub CreateProjectDirectory()
        Try
            Directory.CreateDirectory(Me.DirectoryName)
        Catch ex As Exception
            Throw New Exception
        End Try
    End Sub

    Public Sub CreateFile(ByVal f As String)
        Dim txt As String : txt = vbNullString
        Dim sw As System.IO.StreamWriter

        Try
            sw = New System.IO.StreamWriter(
                f, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
            )
            sw.Write(txt)
        Catch ex As Exception
            '
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
        End Try
    End Sub

    Private Sub CreateProjectModelFile()
        Call CreateFile(Me.ModelFileName)
    End Sub

    Private Sub CreateProjectIniFile()
        Call CreateFile(Me.IniFileName)
    End Sub

    Private Sub CreateProjectInfoFile()
        Call CreateFile(Me.ProjectInfoFileName)
    End Sub

    'Public Overloads Function LoadProject(ByVal p As ProjectInfoModel) As ProjectInfoModel
    '    Dim project As ProjectInfoModel
    '    project = p.ModelLoad(p.ProjectInfoFileName)
    '    LoadProject = project
    'End Function

    'Public Overloads Function LoadProject(ByVal f As String) As ProjectInfoModel
    '    Dim project As ProjectInfoModel
    '    Dim ml As New JsonHandler(Of Object)
    '    project = ml.ModelLoad(Of ProjectInfoModel)(f)
    '    LoadProject = project
    'End Function

    ' この関数はプロジェクトディレクトリの存在チェックを行います
    Public Function CheckStructure() As Integer
        Dim code As Integer : code = -1
        If Directory.Exists(Me.DirectoryName) Then
            If File.Exists(Me.ProjectInfoFileName) Then
                If File.Exists(Me.IniFileName) Then
                    code = 0
                Else
                    code = 10
                End If
            Else
                code = 100
            End If
        Else
            code = 1000
        End If
        CheckStructure = code
    End Function

    ' CheckStructureを真偽値で評価します
    Private Function _CheckProjectExist() As Boolean
        Dim i = -1
        Dim b = False

        i = Me.CheckStructure()
        If i = 0 Then
            b = True
        End If

        _CheckProjectExist = b
    End Function

    Private Function _CheckProjectNotExist() As Boolean
        Return (Not _CheckProjectExist())
    End Function
    Public Function CheckProjectNotExist() As Boolean
        Return _CheckProjectNotExist()
    End Function


    ' ロード可能かのチェック
    Private Function _CheckProjectInfo() As Boolean
        'Dim jh As New JsonHandler(Of Object)
        'Return (jh.CheckModel(Of ProjectInfoModel)(Me.ProjectInfoFileName))
        Return MyBase.CheckModel(Me.ProjectInfoFileName)
    End Function

    Public Function CheckProjectInfo() As Boolean
        Return CheckProjectInfo()
    End Function

    ' ロード可能かのチェック
    Private Function _CheckModel() As Boolean
        Dim jh As New JsonHandler(Of Object)
        Return (jh.CheckModel(Of Model)(Me.ModelFileName))
    End Function
    'Public Function CheckModel() As Boolean
    '    Return _CheckModel()
    'End Function


    ' プロジェクト状態レベル
    Public Function CheckProject() As Integer
        ' 0 ... チェック全てＯＫ
        ' 1 ... ディレクトリが不正
        ' 2 ... プロジェクト情報が不正
        ' 3 ... モデルが不正
        Dim i As Integer : i = -1
        Dim da As New DelegateAction With {
            .CanExecuteHandler = AddressOf _CheckProjectExist,
            .CanExecuteHandler2 = AddressOf _CheckProjectInfo,
            .CanExecuteHandler3 = AddressOf _CheckModel
        }

        i = da.CheckCanExecute(Me)
        CheckProject = i
    End Function

    ' この関数はプロジェクトの作成を行い、その結果を返します
    Public Sub Launch()
        Call Me.CreateProjectDirectory()
        Call Me.CreateProjectIniFile()
        Call Me.CreateProjectInfoFile()
        Call Me.CreateProjectModelFile()
    End Sub

    ' コマンドプロパティ（Ｐｒｏｊｅｃｔをピン止め）
    '---------------------------------------------------------------------------------------------'
    Private _FixProjectCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property FixProjectCommand As ICommand
        Get
            If Me._FixProjectCommand Is Nothing Then
                Me._FixProjectCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _FixProjectCommandExecute,
                    .CanExecuteHandler = AddressOf _FixProjectCommandCanExecute
                }
                Return Me._FixProjectCommand
            Else
                Return Me._FixProjectCommand
            End If
        End Get
    End Property

    Private Sub _CheckFixProjectCommandEnabled()
        Dim b As Boolean : b = True
        Me._FixProjectCommandEnableFlag = b
    End Sub

    Private __FixProjectCommandEnableFlag As Boolean
    Private Property _FixProjectCommandEnableFlag As Boolean
        Get
            Return Me.__FixProjectCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__FixProjectCommandEnableFlag = value
            RaisePropertyChanged("_FixProjectCommandEnableFlag")
            CType(FixProjectCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _FixProjectCommandExecute(ByVal parameter As Object)
        Call DelegateEventListener.Instance.RaiseFixProjectRequested(Me)
    End Sub

    Private Function _FixProjectCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._FixProjectCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    ' コマンドプロパティ（Ｐｒｏｊｅｃｔのピン解除）
    '---------------------------------------------------------------------------------------------'
    Private _RemoveFixedProjectCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property RemoveFixedProjectCommand As ICommand
        Get
            If Me._RemoveFixedProjectCommand Is Nothing Then
                Me._RemoveFixedProjectCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _RemoveFixedProjectCommandExecute,
                    .CanExecuteHandler = AddressOf _RemoveFixedProjectCommandCanExecute
                }
                Return Me._RemoveFixedProjectCommand
            Else
                Return Me._RemoveFixedProjectCommand
            End If
        End Get
    End Property

    Private Sub _CheckRemoveFixedProjectCommandEnabled()
        Dim b As Boolean : b = True
        Me._RemoveFixedProjectCommandEnableFlag = b
    End Sub

    Private __RemoveFixedProjectCommandEnableFlag As Boolean
    Private Property _RemoveFixedProjectCommandEnableFlag As Boolean
        Get
            Return Me.__RemoveFixedProjectCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__RemoveFixedProjectCommandEnableFlag = value
            RaisePropertyChanged("_RemoveFixedProjectCommandEnableFlag")
            CType(RemoveFixedProjectCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _RemoveFixedProjectCommandExecute(ByVal parameter As Object)
        Call DelegateEventListener.Instance.RaiseRemoveFixedProjectRequested(Me)
    End Sub

    Private Function _RemoveFixedProjectCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._RemoveFixedProjectCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    '' コマンドプロパティ（ファイル編集）
    ''---------------------------------------------------------------------------------------------'
    'Private _EditFileCommand As ICommand
    '<JsonIgnore>
    'Public ReadOnly Property EditFileCommand As ICommand
    '    Get
    '        If Me._EditFileCommand Is Nothing Then
    '            Me._EditFileCommand = New DelegateCommand With {
    '                .ExecuteHandler = AddressOf _EditFileCommandExecute,
    '                .CanExecuteHandler = AddressOf _EditFileCommandCanExecute
    '            }
    '            Return Me._EditFileCommand
    '        Else
    '            Return Me._EditFileCommand
    '        End If
    '    End Get
    'End Property

    'Private Sub _CheckEditFileCommandEnabled()
    '    Dim b As Boolean : b = True
    '    Me._EditFileCommandEnableFlag = b
    'End Sub

    'Private __EditFileCommandEnableFlag As Boolean
    'Private Property _EditFileCommandEnableFlag As Boolean
    '    Get
    '        Return Me.__EditFileCommandEnableFlag
    '    End Get
    '    Set(value As Boolean)
    '        Me.__EditFileCommandEnableFlag = value
    '        RaisePropertyChanged("_EditFileCommandEnableFlag")
    '        CType(EditFileCommand, DelegateCommand).RaiseCanExecuteChanged()
    '    End Set
    'End Property

    'Private Sub _EditFileCommandExecute(ByVal parameter As Object)
    '    Dim f = CType(parameter, String)
    '    Dim psi As New System.Diagnostics.ProcessStartInfo With {
    '        .FileName = "notepad.exe",
    '        .Arguments = $"{f}"
    '    }
    '    Dim p = System.Diagnostics.Process.Start(psi)
    'End Sub

    'Private Function _EditFileCommandCanExecute(ByVal parameter As Object) As Boolean
    '    Return Me._EditFileCommandEnableFlag
    'End Function
    ''---------------------------------------------------------------------------------------------'

    '' コマンドプロパティ（ファイル選択）
    ''---------------------------------------------------------------------------------------------'
    'Private _InputFileDialogCommand As ICommand
    '<JsonIgnore>
    'Public ReadOnly Property InputFileDialogCommand As ICommand
    '    Get
    '        If Me._InputFileDialogCommand Is Nothing Then
    '            Me._InputFileDialogCommand = New DelegateCommand With {
    '                .ExecuteHandler = AddressOf _InputFileDialogCommandExecute,
    '                .CanExecuteHandler = AddressOf _InputFileDialogCommandCanExecute
    '            }
    '            Return Me._InputFileDialogCommand
    '        Else
    '            Return Me._InputFileDialogCommand
    '        End If
    '    End Get
    'End Property

    'Private Sub _CheckInputFileDialogCommandEnabled()
    '    Dim b As Boolean : b = True
    '    Me._InputFileDialogCommandEnableFlag = b
    'End Sub

    'Private __InputFileDialogCommandEnableFlag As Boolean
    'Private Property _InputFileDialogCommandEnableFlag As Boolean
    '    Get
    '        Return Me.__InputFileDialogCommandEnableFlag
    '    End Get
    '    Set(value As Boolean)
    '        Me.__InputFileDialogCommandEnableFlag = value
    '        RaisePropertyChanged("_InputFileDialogCommandEnableFlag")
    '        CType(InputFileDialogCommand, DelegateCommand).RaiseCanExecuteChanged()
    '    End Set
    'End Property

    'Private Sub _InputFileDialogCommandExecute(ByVal parameter As Object)
    '    Dim f As String
    '    Dim t As String
    '    Dim fil = "JSONファイル(*.json)|*.json"
    '    Select Case parameter
    '        Case Me.ModelFileName
    '            f = "Model.json"
    '            t = "Select Model File"
    '        Case Me.ViewModelFileName
    '            f = "ViewModel.json"
    '            t = "Select ViewModel File"
    '        Case Me.ProjectInfoFileName
    '            f = "ProjectInfo.json"
    '            t = "Select Project Information File"
    '        Case Me.IconFileName
    '            f = "Project.ico"
    '            t = "プロジェクトアイコンの選択"
    '            fil = "すべてのファイル(*.*)|*.*"
    '    End Select
    '    Dim ofd As New OpenFileDialog With {
    '        .FileName = f,
    '        .InitialDirectory = Me.DirectoryName,
    '        .Filter = fil,
    '        .FilterIndex = 1,
    '        .Title = t
    '    }
    '    If ofd.ShowDialog() = DialogResult.OK Then
    '        Me.ModelFileName = ofd.FileName
    '    End If
    'End Sub

    'Private Function _InputFileDialogCommandCanExecute(ByVal parameter As Object) As Boolean
    '    Return Me._InputFileDialogCommandEnableFlag
    'End Function
    ''---------------------------------------------------------------------------------------------'

    '' コマンドプロパティ（フォルダ選択）
    ''---------------------------------------------------------------------------------------------'
    'Private _FolderBrowserDialogCommand As ICommand
    '<JsonIgnore>
    'Public ReadOnly Property FolderBrowserDialogCommand As ICommand
    '    Get
    '        If Me._FolderBrowserDialogCommand Is Nothing Then
    '            Me._FolderBrowserDialogCommand = New DelegateCommand With {
    '                .ExecuteHandler = AddressOf _FolderBrowserDialogCommandExecute,
    '                .CanExecuteHandler = AddressOf _FolderBrowserDialogCommandCanExecute
    '            }
    '            Return Me._FolderBrowserDialogCommand
    '        Else
    '            Return Me._FolderBrowserDialogCommand
    '        End If
    '    End Get
    'End Property

    'Private Sub _CheckFolderBrowserDialogCommandEnabled()
    '    Dim b As Boolean : b = True
    '    Me._FolderBrowserDialogCommandEnableFlag = b
    'End Sub

    'Private __FolderBrowserDialogCommandEnableFlag As Boolean
    'Private Property _FolderBrowserDialogCommandEnableFlag As Boolean
    '    Get
    '        Return Me.__FolderBrowserDialogCommandEnableFlag
    '    End Get
    '    Set(value As Boolean)
    '        Me.__FolderBrowserDialogCommandEnableFlag = value
    '        RaisePropertyChanged("_FolderBrowserDialogCommandEnableFlag")
    '        CType(FolderBrowserDialogCommand, DelegateCommand).RaiseCanExecuteChanged()
    '    End Set
    'End Property

    'Private Sub _FolderBrowserDialogCommandExecute(ByVal parameter As Object)
    '    Dim f = CType(parameter, String)
    '    Dim fbd As New FolderBrowserDialog With {
    '        .Description = "プロジェクトフォルダの指定",
    '        .RootFolder = Me.RootDirectoryName,
    '        .SelectedPath = Me.DirectoryName,
    '        .ShowNewFolderButton = True
    '    }
    '    If fbd.ShowDialog() = DialogResult.OK Then
    '        Me.DirectoryName = fbd.SelectedPath
    '    End If
    'End Sub

    'Private Function _FolderBrowserDialogCommandCanExecute(ByVal parameter As Object) As Boolean
    '    Return Me._FolderBrowserDialogCommandEnableFlag
    'End Function
    ''---------------------------------------------------------------------------------------------'

    Sub New()
        Call _CheckFixProjectCommandEnabled()
        Call _CheckRemoveFixedProjectCommandEnabled()
    End Sub
End Class
