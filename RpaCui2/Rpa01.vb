Imports System.IO

Public Class Rpa01 : Inherits RpaBase(Of Rpa01)
    Private Const EXCEL_APP As String = "Excel.Application"
    Private Const CSCRIPT As String = "cscript"

    Private _PasswordOfAttacheCase As String
    Public Property PasswordOfAttacheCase As String
        Get
            Return Me._PasswordOfAttacheCase
        End Get
        Set(value As String)
            Me._PasswordOfAttacheCase = value
        End Set
    End Property

    Private _OutlookSourceInboxName As String
    Public Property OutlookSourceInboxName As String
        Get
            Return Me._OutlookSourceInboxName
        End Get
        Set(value As String)
            Me._OutlookSourceInboxName = value
        End Set
    End Property

    Private _OutlookBackupInboxName As String
    Public Property OutlookBackupInboxName As String
        Get
            Return Me._OutlookBackupInboxName
        End Get
        Set(value As String)
            Me._OutlookBackupInboxName = value
        End Set
    End Property

    Private _MasterCsvFileName As String
    Public Property MasterCsvFileName As String
        Get
            Return Me._MasterCsvFileName
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Me._MasterCsvFileName = "master.csv"
            Else
                Me._MasterCsvFileName = value
            End If
        End Set
    End Property

    Private _RestartCode As String
    Public Property RestartCode As String
        Get
            Return Me._RestartCode
        End Get
        Set(value As String)
            Me._RestartCode = value
        End Set
    End Property

    Public Overrides Function SetupProjectObject(project As String) As Object
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Main(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer

        Dim master = rpa.MyProjectDirectory & "\" & Me.MasterCsvFileName
        Console.WriteLine("以下のファイルを用意したら、[Enter]キーをクリックしてください")
        Console.WriteLine("ファイル名 : " & Me.MasterCsvFileName)
        Console.ReadLine()
        If Not File.Exists(master) Then
            Console.WriteLine("エラー：ファイルが存在しません : " & master)
            Return 1000
        End If

        ' ワークディレクトリの作成
        Dim wd = rpa.MyProjectDirectory & "\work"
        If Not Directory.Exists(wd) Then
            Directory.CreateDirectory(wd)
        End If

        Dim wmaster = wd & "\" & Me.MasterCsvFileName
        File.Copy(master, wmaster, True)

        Console.WriteLine("添付ファイルを検索しています..."

        Dim wtemp = wd & "\tmp.csv"
    End Function
End Class
