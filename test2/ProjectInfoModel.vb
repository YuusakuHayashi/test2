Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ProjectInfoModel
    Inherits JsonHandler(Of ProjectInfoModel)

    Private Const SHIFT_JIS As String = "Shift-JIS"

    Private _ImageFileName As String
    Public Property ImageFileName As String
        Get
            Return _ImageFileName
        End Get
        Set(value As String)
            _ImageFileName = value
        End Set
    End Property

    Private _Image As BitmapImage
    <JsonIgnore>
    Public Property [Image] As BitmapImage
        Get
            Return _Image
        End Get
        Set(value As BitmapImage)
            _Image = value
        End Set
    End Property

    Private _Index As Integer
    Public Property [Index] As String
        Get
            Return _Index
        End Get
        Set(value As String)
            _Index = value
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

    Private _DirectoryName As String
    Public Property DirectoryName As String
        Get
            Return _DirectoryName
        End Get
        Set(value As String)
            _DirectoryName = value
        End Set
    End Property

    ' プロジェクトのモデルファイルです
    Public ReadOnly Property ModelFileName As String
        Get
            Return DirectoryName & "\Model.json"
        End Get
    End Property

    ' プロジェクトがアプリケーションにより作成されたことを表すファイルです
    Public ReadOnly Property IniFileName As String
        Get
            Return DirectoryName & "\Project.ini"
        End Get
    End Property

    Public ReadOnly Property ProjectInfoFileName As String
        Get
            Return DirectoryName & "\ProjectInfo.json"
        End Get
    End Property

    Public ReadOnly Property ViewModelFileName As String
        Get
            Return DirectoryName & "\ViewModel.json"
        End Get
    End Property

    Public Delegate Sub ProjectLaunchProxy()

    Public Sub ProjectSave()
        Me.ModelSave(Me.ProjectInfoFileName, Me)
    End Sub

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

    Public Overloads Function LoadProject(ByVal p As ProjectInfoModel) As ProjectInfoModel
        Dim project As ProjectInfoModel
        project = p.ModelLoad(p.ProjectInfoFileName)
        LoadProject = project
    End Function

    Public Overloads Function LoadProject(ByVal f As String) As ProjectInfoModel
        Dim project As ProjectInfoModel
        Dim ml As New JsonHandler(Of Object)
        project = ml.ModelLoad(Of ProjectInfoModel)(f)
        LoadProject = project
    End Function

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
End Class
