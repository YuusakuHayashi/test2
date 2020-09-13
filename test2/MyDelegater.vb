Imports System.Data
Imports System.Data.SqlClient

Delegate Function ModelLoadProxy1(Of T)(ByVal f As String) As T
Delegate Function ModelLoadProxy2(Of T)(ByVal f As String, ByVal k As String) As T
Delegate Sub ModelSaveProxy1(Of T)(ByVal f As String, ByVal m As T)
Delegate Sub ModelSaveProxy2(Of T, T2)(ByVal f As String, ByVal m As T, ByVal key As String)
Delegate Sub ProjectProxy(ByVal f As String)
Delegate Sub ProjectProxy2(ByVal f As String, ByVal txt As String)
Delegate Function SaveHandler(Of T)(ByVal f As String, ByVal m As T)
Public Delegate Function GetDataSetProxy(ByVal scmd As SqlCommand) As DataSet
Public Delegate Function LoadHandler(Of T)(ByVal f As String) As T

Module MyDelegater
    Public Sub ViewFlagCheck(ByVal obj As Object)
    End Sub
End Module
