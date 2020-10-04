Imports System.IO

Public Class ProjectInfoModel
    Inherits ModelLoader(Of ProjectInfoModel)

    Private Const SHIFT_JIS As String = "Shift-JIS"

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

    ' プロジェクトのビューモデルファイルです
    Public ReadOnly Property ViewModelFileName As String
        Get
            Return DirectoryName & "\ViewModel.json"
        End Get
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

    Public Delegate Sub ProjectLaunchProxy()

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

    'Private Sub CreateProjectModelFile()
    '    Try
    '        File.Create(Me.ModelFileName)
    '    Catch ex As Exception
    '        Throw New Exception
    '    End Try
    'End Sub

    'Private Sub CreateProjectViewModelFile()
    '    Try
    '        File.Create(Me.ViewModelFileName)
    '    Catch ex As Exception
    '        Throw New Exception
    '    End Try
    'End Sub

    'Private Sub CreateProjectIniFile()
    '    Try
    '        File.Create(Me.IniFileName)
    '    Catch ex As Exception
    '        Throw New Exception
    '    End Try
    'End Sub

    Private Sub CreateProjectModelFile()
        Call CreateFile(Me.ModelFileName)
    End Sub

    Private Sub CreateProjectViewModelFile()
        Call CreateFile(Me.ViewModelFileName)
    End Sub

    Private Sub CreateProjectIniFile()
        Call CreateFile(Me.IniFileName)
    End Sub


    ' この関数はプロジェクトディレクトリの存在チェックを行います
    Public Function CheckProjectDirectory() As Integer
        Dim code As Integer : code = 99
        If Directory.Exists(Me.DirectoryName) Then
            code = 10
            If File.Exists(Me.IniFileName) Then
                code = 0
            End If
        End If
        CheckProjectDirectory = code
    End Function


    ' この関数はプロジェクトの作成を行い、その結果を返します
    Public Function ProjectLaunch() As Integer
        Dim proxy(3) As ProjectLaunchProxy
        Dim proxy2 As ProjectLaunchProxy

        proxy(0) = AddressOf CreateProjectDirectory
        proxy(1) = AddressOf CreateProjectIniFile
        proxy(2) = AddressOf CreateProjectModelFile
        proxy(3) = AddressOf CreateProjectViewModelFile

        Dim code As Integer : code = CheckProjectDirectory()
        Try
            Select Case code
                Case 0
                Case 10
                Case 99
                    proxy2 = [Delegate].Combine(proxy)
                Case Else
            End Select
            If proxy2 IsNot Nothing Then
                If proxy2.GetInvocationList IsNot Nothing Then
                    Call proxy2()
                    code = 0
                End If
            End If
        Catch ex As Exception
        Finally
            ProjectLaunch = code
        End Try
    End Function
End Class
