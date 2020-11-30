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
            Me._MasterCsvFileName = value
        End Set
    End Property

    Public Overrides Function SetupProjectObject(project As String) As Object
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Main(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        If String.IsNullOrEmpty(Me.MasterCsvFileName) Then
            Me.MasterCsvFileName = rpa.MyProjectDirectory & "\" & "master.csv"
        End If
        Console.WriteLine("以下のファイルを用意したら、[Enter]キーをクリックしてください")
        Console.WriteLine("ファイル名 : " & Me.MasterCsvFileName)
    End Function
End Class
