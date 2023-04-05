Imports System.Data.SqlClient
Public Class ApplicationMain
    Public Shared oReturnStatus As New ReturnStatus
    Public Shared StrSql As String = String.Empty
    Public Shared StrSql2 As String = String.Empty
    Public Shared oSqlConnectionStringBuilder As New SqlConnectionStringBuilder
    Public Shared oSqlConnection As New SqlConnection
    Public Shared oSqlCommand As New SqlCommand
    Public Shared oSqlTransaction As SqlTransaction
    Public Shared oSqlDataAdapter As New SqlDataAdapter
    Public Shared oSQLAction As New SQLAction
    'Public Shared ocontrols As New Controls
    'Public Shared oclsMemory As New clsMemory
    Sub New()
        Try
            oSqlConnection = OpenNewConnection()
            oSqlConnection.Open()
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Sub New(_SqlConnection As SqlConnection)
        Try
            oSqlConnection = _SqlConnection
            If oSqlConnection.State = ConnectionState.Closed Then oSqlConnection.Open()
        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class
