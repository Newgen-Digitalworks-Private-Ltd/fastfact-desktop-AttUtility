Imports System.Data.SqlClient
Imports AttendanceUtility.ApplicationMain
Public Class SQLAction
    Public Function ExecuteDataTables(sqlQuery As String, sqlCmdType As CommandType, Optional ParaValue As List(Of SqlParameter) = Nothing) As ReturnStatus
        Try
            Dim dt As New DataTable
            oSqlCommand = New SqlCommand
            oSqlCommand.CommandText = sqlQuery
            oSqlCommand.CommandType = sqlCmdType
            oSqlCommand.Connection = oSqlConnection
            oSqlCommand.CommandTimeout = 0
            oSqlCommand.Transaction = oSqlTransaction
            If Not IsNothing(ParaValue) Then
                oSqlCommand.Parameters.AddRange(ParaValue.ToArray)
            End If
            oSqlDataAdapter.SelectCommand = oSqlCommand
            oSqlDataAdapter.Fill(dt)
            Return New ReturnStatus(True, dt)
        Catch ex As Exception
            Return New ReturnStatus(False, ex.Message)
        Finally
            'If oSqlConnection.State = ConnectionState.Open Then oSqlConnection.Close() : oSqlConnection.Open()
        End Try
    End Function
    Public Function ExecuteScalars(sqlQuery As String, sqlCmdType As CommandType) As ReturnStatus
        Try
            oSqlCommand = New SqlCommand
            oSqlCommand.Connection = oSqlConnection
            oSqlCommand.CommandTimeout = 0
            oSqlCommand.CommandText = sqlQuery
            oSqlCommand.CommandType = sqlCmdType
            oSqlCommand.Transaction = oSqlTransaction
            Return New ReturnStatus(True, oSqlCommand.ExecuteScalar)
        Catch ex As Exception
            Return New ReturnStatus(False, ex.Message)
        Finally
            'If oSqlConnection.State = ConnectionState.Open Then oSqlConnection.Close() : oSqlConnection.Open()
        End Try
    End Function
    Public Function ExecuteNonQueryscmd(sqlQuery As String, sqlCmdType As CommandType, Optional ParaValue As List(Of SqlParameter) = Nothing) As ReturnStatus
        Try
            oSqlCommand = New SqlCommand
            oSqlCommand.CommandText = sqlQuery
            oSqlCommand.CommandType = sqlCmdType
            oSqlCommand.Connection = oSqlConnection
            oSqlCommand.CommandTimeout = 0
            oSqlCommand.Transaction = oSqlTransaction
            If Not IsNothing(ParaValue) Then
                oSqlCommand.Parameters.AddRange(ParaValue.ToArray)
            End If
            oSqlCommand.ExecuteNonQuery()
            Return New ReturnStatus(True)
        Catch ex As Exception
            Return New ReturnStatus(False, ex.Message)
        Finally
            'If oSqlConnection.State = ConnectionState.Open Then oSqlConnection.Close() : oSqlConnection.Open()
        End Try
    End Function
    Public Function TranBegin() As ReturnStatus
        Try
            oSqlCommand.Connection = oSqlConnection
            oSqlTransaction = oSqlCommand.Connection.BeginTransaction()
            oSqlCommand.Transaction = oSqlTransaction
            Return New ReturnStatus(True)
        Catch ex As Exception
            Return New ReturnStatus(False, ex.Message)
        Finally
            'If oSqlConnection.State = ConnectionState.Open Then oSqlConnection.Close() : oSqlConnection.Open()
        End Try
    End Function
    Public Function TranCommit() As ReturnStatus
        Try
            oSqlCommand.Connection = oSqlConnection
            oSqlCommand.Transaction.Commit()
            oSqlTransaction = Nothing
            Return New ReturnStatus(True)
        Catch ex As Exception
            Return New ReturnStatus(False, ex.Message)
        Finally
            'If oSqlConnection.State = ConnectionState.Open Then oSqlConnection.Close() : oSqlConnection.Open()
        End Try
    End Function
    Public Function TranRollBack() As ReturnStatus
        Try
            oSqlCommand.Connection = oSqlConnection
            oSqlCommand.Transaction.Rollback()
            oSqlTransaction = Nothing
            Return New ReturnStatus(True)
        Catch ex As Exception
            Return New ReturnStatus(False, ex.Message)
        Finally
            'If oSqlConnection.State = ConnectionState.Open Then oSqlConnection.Close() : oSqlConnection.Open()
        End Try
    End Function

    Public Function ExecuteNonQuerySPscmd(sqlQuery As String, sqlCmdType As CommandType, dt As DataTable, strParamsName As String) As ReturnStatus
        Try
            oSqlCommand = New SqlCommand
            oSqlCommand.CommandText = sqlQuery
            oSqlCommand.CommandType = sqlCmdType
            oSqlCommand.Connection = oSqlConnection
            oSqlCommand.CommandTimeout = 0
            oSqlCommand.Transaction = oSqlTransaction
            If Not IsNothing(dt) Then
                oSqlCommand.Parameters.AddWithValue(strParamsName, dt)
            End If
            oSqlCommand.ExecuteNonQuery()
            Return New ReturnStatus(True)
        Catch ex As Exception
            Return New ReturnStatus(False, ex.Message)
        Finally
            'If oSqlConnection.State = ConnectionState.Open Then oSqlConnection.Close() : oSqlConnection.Open()
        End Try
    End Function
End Class
