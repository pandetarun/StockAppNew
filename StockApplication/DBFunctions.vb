Imports FirebirdSql.Data.FirebirdClient
Imports System.IO.StreamReader
Imports System.Globalization
Imports System.Data
Imports System.IO
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class DBFunctions

    Public Shared myConnection As New FbConnection
    Public Shared myConnectionforDataCollection As New FbConnection
    Public Shared myConnectionforCalculation As New FbConnection
    Public Shared myConnectionforIndication As New FbConnection
    Public Shared myDataSet As New DataSet
    Public Shared ServerType As FbServerType = FbServerType.Default
    Public Shared Database As String = My.Settings.ApplicationFileLocation & My.Settings.DataBase
    Public Shared DataSource As String = My.Settings.DataSource
    Public Shared Password As String = My.Settings.Password
    Public Shared UserID As String = My.Settings.UserID

    Public Shared Function OpenSQLConnection() As Boolean

        Try
            If myConnection.State = ConnectionState.Closed Then
                StockAppLogger.Log("OpenSQLConnection Start", "DBFunctions")
                Dim cs As New FbConnectionStringBuilder()
                cs.DataSource = DataSource
                cs.Password = Password
                cs.UserID = UserID
                cs.Port = 3050
                cs.Pooling = False
                cs.Database = Database
                cs.Charset = "UNICODE_FSS"
                cs.ConnectionLifeTime = 30
                cs.ServerType = ServerType
                myDataSet.Locale = CultureInfo.InvariantCulture
                myConnection.ConnectionString = cs.ToString
                If cs IsNot Nothing Then cs = Nothing
                myConnection.Open()
                myDataSet.Reset()
                StockAppLogger.Log("OpenSQLConnection End", "DBFunctions")
            End If
            Return True
        Catch ex As Exception
            StockAppLogger.LogError("OpenSQLConnection Error Occurred in opening the connection ", ex, "DBFunctions")
            Return False
        End Try
    End Function

    Public Shared Function OpenSQLConnection(tmpmyConnection As FbConnection) As Boolean

        Try
            If tmpmyConnection.State = ConnectionState.Closed Then
                StockAppLogger.Log("OpenSQLConnection Start", "DBFunctions")
                Dim cs As New FbConnectionStringBuilder()
                cs.DataSource = DataSource
                cs.Password = Password
                cs.UserID = UserID
                cs.Port = 3050
                cs.Pooling = False
                cs.Database = Database
                cs.Charset = "UNICODE_FSS"
                cs.ConnectionLifeTime = 30
                cs.ServerType = ServerType
                myDataSet.Locale = CultureInfo.InvariantCulture
                tmpmyConnection.ConnectionString = cs.ToString
                If cs IsNot Nothing Then cs = Nothing
                tmpmyConnection.Open()
                myDataSet.Reset()
                StockAppLogger.Log("OpenSQLConnection End", "DBFunctions")
            End If
            Return True
        Catch ex As Exception
            StockAppLogger.LogError("OpenSQLConnection Error Occurred in opening the connection ", ex, "DBFunctions")
            Return False
        End Try
    End Function

    Public Shared Function CreateDatabase() As Boolean
        StockAppLogger.Log("CreateDatabase Start", "DBFunctions")
        Try
            Dim cs = New FbConnectionStringBuilder()
            'If Not ServerType = FbServerType.Default Then
            cs.DataSource = DataSource
            cs.Password = Password
            cs.UserID = UserID
            cs.Port = 3050
            'End If

            cs.Pooling = False
            cs.Database = Database
            cs.Charset = "UNICODE_FSS"
            cs.ServerType = ServerType
            FbConnection.CreateDatabase(cs.ToString)
            If cs IsNot Nothing Then cs = Nothing
            StockAppLogger.Log("CreateDatabase End", "DBFunctions")
            Return True
        Catch ex As Exception
            StockAppLogger.Log("CreateDatabase Error in creating database", "DBFunctions")
            Return False
        End Try
    End Function

    Public Shared Sub CloseSQLConnection()
        If myConnection.State = ConnectionState.Open Then
            myConnection.Close()
        End If
    End Sub

    Public Shared Sub CloseSQLConnectionExt(ByVal transactionType As String)
        If transactionType IsNot Nothing And transactionType = "DC" Then
            If myConnectionforDataCollection.State = ConnectionState.Open Then
                myConnectionforDataCollection.Close()
            End If
        ElseIf transactionType IsNot Nothing And transactionType = "CI" Then
            If myConnectionforCalculation.State = ConnectionState.Open Then
                myConnectionforCalculation.Close()
            End If
        ElseIf transactionType IsNot Nothing And transactionType = "DI" Then
            If myConnectionforIndication.State = ConnectionState.Open Then
                myConnectionforIndication.Close()
            End If
        End If
    End Sub

    Public Shared Function ExecuteSQLStmt(ByVal sSQL As String, Optional ByVal Disconnect As Boolean = False) As Boolean
        StockAppLogger.Log("ExecuteSQLStmt Start", "DBFunctions")
        If OpenSQLConnection() = True Then
            Dim myCmd As New FbCommand(sSQL, myConnection)

            Try
                myCmd.ExecuteNonQuery()
                StockAppLogger.Log("ExecuteSQLStmt End", "DBFunctions")
                Return True
            Catch ex As Exception
                StockAppLogger.LogError("ExecuteSQLStmtError in executing statement" & sSQL, ex, "DBFunctions")
                Return False
            Finally
                If Disconnect = True Then CloseSQLConnection()
                If myCmd IsNot Nothing Then
                    myCmd.Dispose()
                    myCmd = Nothing
                End If
            End Try
        Else
            Return False
        End If
    End Function

    Public Shared Function ExecuteSQLStmtExt(ByVal sSQL As String, ByVal transactionType As String, Optional ByVal Disconnect As Boolean = False) As Boolean
        StockAppLogger.Log("ExecuteSQLStmt Start", "DBFunctions")
        'DC for Data Collect
        'CI for Calculate Indicators
        'DI for Define Indication

        If transactionType IsNot Nothing And transactionType = "DC" Then
            If OpenSQLConnection(myConnectionforDataCollection) = False Then
                Return False
            End If
        ElseIf transactionType IsNot Nothing And transactionType = "CI" Then
            If OpenSQLConnection(myConnectionforCalculation) = False Then
                Return False
            End If
        ElseIf transactionType IsNot Nothing And transactionType = "DI" Then
            If OpenSQLConnection(myConnectionforIndication) = False Then
                Return False
            End If
        End If
        'If OpenSQLConnection() = True Then
        Dim myCmd As New FbCommand(sSQL, myConnection)

        Try
            myCmd.ExecuteNonQuery()
            StockAppLogger.Log("ExecuteSQLStmt End", "DBFunctions")
            Return True
        Catch ex As Exception
            StockAppLogger.LogError("ExecuteSQLStmtError in executing statement" & sSQL, ex, "DBFunctions")
            Return False
        Finally
            If Disconnect = True Then CloseSQLConnection()
            If myCmd IsNot Nothing Then
                myCmd.Dispose()
                myCmd = Nothing
            End If
        End Try
        'Else
        '    Return False
        'End If
    End Function

    Public Shared Function ExecuteSQLStmtandReturnResult(ByVal sSQL As String, Optional ByVal Disconnect As Boolean = False) As FbDataReader
        StockAppLogger.Log("ExecuteSQLStmt Start", "DBFunctions")
        Dim conn As New FbConnection
        Dim command As New FbCommand
        Dim ds As FbDataReader = Nothing
        Dim sql As String = ""
        'Dim resultList As List(Of String)

        StockAppLogger.Log("getDataFromTable Start", "DBFunctions")

        If OpenSQLConnection() = True Then
            Try
                'excecuted the SQL command 

                command.Connection = myConnection
                command.CommandText = sSQL
                ds = command.ExecuteReader
                'resultList = New List(Of String)
                StockAppLogger.Log("getDataFromTable End", "DBFunctions")
                Return ds
            Catch ex As Exception
                StockAppLogger.LogError("getDataFromTable Error in getting data for query " & sql, ex, "DBFunctions")
                Return Nothing
            End Try
        End If
        Return Nothing
    End Function

    Public Shared Function ExecuteSQLStmtandReturnResultExt(ByVal sSQL As String, Optional ByVal transactionType As String = "DC", Optional ByVal Disconnect As Boolean = False) As FbDataReader
        StockAppLogger.Log("ExecuteSQLStmt Start", "DBFunctions")
        Dim conn As New FbConnection
        Dim command As New FbCommand
        Dim ds As FbDataReader = Nothing
        Dim sql As String = ""
        'Dim resultList As List(Of String)

        StockAppLogger.Log("getDataFromTable Start", "DBFunctions")
        'DC for Data Collect
        'CI for Calculate Indicators
        'DI for Define Indication

        If transactionType IsNot Nothing And transactionType = "DC" Then
            If OpenSQLConnection(myConnectionforDataCollection) = False Then
                Return Nothing
            End If
        ElseIf transactionType IsNot Nothing And transactionType = "CI" Then
            If OpenSQLConnection(myConnectionforCalculation) = False Then
                Return Nothing
            End If
        ElseIf transactionType IsNot Nothing And transactionType = "DI" Then
            If OpenSQLConnection(myConnectionforIndication) = False Then
                Return Nothing
            End If
        End If
        'If OpenSQLConnection() = True Then
        Try
            'excecuted the SQL command 

            command.Connection = myConnection
            command.CommandText = sSQL
            ds = command.ExecuteReader
            'resultList = New List(Of String)
            StockAppLogger.Log("getDataFromTable End", "DBFunctions")
            Return ds
        Catch ex As Exception
            StockAppLogger.LogError("getDataFromTable Error in getting data for query " & sql, ex, "DBFunctions")
            Return Nothing
        End Try
        ' End If
        Return Nothing
    End Function

    Public Shared Function getDataFromTable(ByVal tableName As String, Optional ByVal columnNames As String = "", Optional ByVal whereClause As String = "", Optional ByVal orderClause As String = "", Optional ByVal Disconnect As Boolean = False) As FbDataReader

        Dim conn As New FbConnection
        Dim command As New FbCommand
        Dim ds As FbDataReader = Nothing
        Dim sql As String = ""
        'Dim resultList As List(Of String)

        StockAppLogger.Log("getDataFromTable Start", "DBFunctions")
        If OpenSQLConnection() = True Then
            Try
                'excecuted the SQL command 
                If columnNames IsNot "" Then
                    sql = "Select " & columnNames & " from " & tableName
                Else
                    sql = "Select * from " & tableName
                End If
                If whereClause IsNot "" Then
                    sql = sql & " where " & whereClause
                End If
                If orderClause IsNot "" Then
                    sql = sql & " order by " & orderClause
                End If
                command.Connection = myConnection
                command.CommandText = sql
                ds = command.ExecuteReader
                'resultList = New List(Of String)
                StockAppLogger.Log("getDataFromTable End", "DBFunctions")
                Return ds
            Catch ex As Exception
                StockAppLogger.LogError("getDataFromTable Error in getting data for query " & sql, ex, "DBFunctions")
                Return Nothing
            End Try
        End If
        Return ds
    End Function


    Public Shared Function getDataFromTableExt(ByVal tableName As String, ByVal transactionType As String, Optional ByVal columnNames As String = "", Optional ByVal whereClause As String = "", Optional ByVal orderClause As String = "", Optional ByVal Disconnect As Boolean = False) As FbDataReader

        Dim conn As New FbConnection
        Dim command As New FbCommand
        Dim ds As FbDataReader = Nothing
        Dim sql As String = ""
        'Dim resultList As List(Of String)

        StockAppLogger.Log("getDataFromTable Start", "DBFunctions")
        'DC for Data Collect
        'CI for Calculate Indicators
        'DI for Define Indication

        If transactionType IsNot Nothing And transactionType = "DC" Then
            If OpenSQLConnection(myConnectionforDataCollection) = False Then
                Return Nothing
            End If
        ElseIf transactionType IsNot Nothing And transactionType = "CI" Then
            If OpenSQLConnection(myConnectionforCalculation) = False Then
                Return Nothing
            End If
        ElseIf transactionType IsNot Nothing And transactionType = "DI" Then
            If OpenSQLConnection(myConnectionforIndication) = False Then
                Return Nothing
            End If
        End If
        'If OpenSQLConnection() = True Then
        Try
            'excecuted the SQL command 
            If columnNames IsNot "" Then
                sql = "Select " & columnNames & " from " & tableName
            Else
                sql = "Select * from " & tableName
            End If
            If whereClause IsNot "" Then
                sql = sql & " where " & whereClause
            End If
            If orderClause IsNot "" Then
                sql = sql & " order by " & orderClause
            End If
            command.Connection = myConnection
            command.CommandText = sql
            ds = command.ExecuteReader
            'resultList = New List(Of String)
            StockAppLogger.Log("getDataFromTable End", "DBFunctions")
            Return ds
        Catch ex As Exception
            StockAppLogger.LogError("getDataFromTable Error in getting data for query " & sql, ex, "DBFunctions")
            Return Nothing
        End Try
        'End If
        Return ds
    End Function
End Class
