Imports System.Runtime.InteropServices

Public Class RpaOutlookUtility : Inherits RpaUtilityBase
    Private _InboxFolderName As String

    Public Delegate Function OutlookAutomation(ByRef ns As Object) As Object


    ' _IsInboxFolderExist(ns)は委譲者で、引数を渡すのが難しいので、専用プロパティを持たせる
    Public Function IsInboxFolderExist(ByVal [folder] As String) As Boolean
        Me._InboxFolderName = vbNullString
        Me._InboxFolderName = [folder]
        Return (CallOutlookMacro(AddressOf _IsInboxFolderExist))
    End Function


    Private Function _IsInboxFolderExist(ByRef ns As Object) As Object
        Dim inbox, inboxfolders
        Dim rtn = False
        Try
            inbox = ns.GetDefaultFolder(6)
            Try
                inboxfolders = inbox.Folders
                For Each inboxfolder In inboxfolders
                    If inboxfolder.Name = Me._InboxFolderName Then
                        rtn = True
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Console.WriteLine($"フォルダーコレクションの取得に失敗しました")
                rtn = False
            Finally
                Marshal.ReleaseComObject(inboxfolders)
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine($"受信トレイの取得に失敗しました")
            rtn = False
        Finally
            Marshal.ReleaseComObject(inbox)
        End Try
        Return rtn
    End Function


    Public Function InvokeOutlookMacroFunction(ByRef automation As OutlookAutomation) As Object
        Dim obj = CallOutlookMacro(automation)
        Return obj
    End Function

    Private Function CallOutlookMacro(ByRef automation As OutlookAutomation) As Object
        Const MAPI = "MAPI"
        Dim olapp, ns
        Dim rtn = Nothing
        Try
            olapp = CreateObject("Outlook.Application")
            Try
                ns = olapp.GetNamespace(MAPI)
                rtn = automation(ns)
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Console.WriteLine($"NameSpaceオブジェクトの取得に失敗しました")
                rtn = Nothing
            Finally
                Marshal.ReleaseComObject(ns)
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine($"Outlook.Applicationオブジェクトの取得に失敗しました")
            rtn = Nothing
        Finally
            Marshal.ReleaseComObject(olapp)
        End Try
        Return rtn
    End Function
End Class
