'Imports FirebirdSql.Data.Firebird
Imports FirebirdSql.Data.FirebirdClient
Imports System.IO.StreamReader
Imports System.Globalization
Imports System.Data
Imports System.IO
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class DBFunctions

    'Public myCellCollection As New Collection
    Public Shared myConnection As New FbConnection
    'Public myDataAdapter() As FbDataAdapter
    Public Shared myDataSet As New DataSet
    'Public myWorkRow As DataRow
    Public Shared ServerType As FbServerType = FbServerType.Default   '!
    Public Shared Database As String = My.Settings.ApplicationFileLocation & "\DB\STOCKAPPDB.fdb"                   '!
    Public Shared DataSource As String = My.Settings.DataSource                            '!
    Public Shared Password As String = My.Settings.Password                              '!
    Public Shared UserID As String = My.Settings.UserID                                '!

    Public Shared Function OpenSQLConnection() As Boolean
        Try
            If myConnection.State = ConnectionState.Closed Then
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
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("An error has occured!" & vbCrLf & vbCrLf &
            ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    Public Function CreateDatabase() As Boolean
        Try
            Dim cs = New FbConnectionStringBuilder()
            If Not ServerType = FbServerType.Default Then
                cs.DataSource = DataSource
                cs.Password = Password
                cs.UserID = UserID
                cs.Port = 3050
            End If
            cs.Pooling = False
            cs.Database = Database
            cs.Charset = "UNICODE_FSS"
            cs.ServerType = ServerType

            FbConnection.CreateDatabase(cs.ToString)
            If cs IsNot Nothing Then cs = Nothing
            Return True
        Catch ex As Exception
            MessageBox.Show("An error has occured!" & vbCrLf & vbCrLf &
            ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Public Shared Sub CloseSQLConnection()

        If myConnection.State = ConnectionState.Open Then
            myConnection.Close()
        End If

    End Sub

    Public Shared Function ExecuteSQLStmt(ByVal sSQL As String, Optional ByVal Disconnect As Boolean = False) As Boolean

        If OpenSQLConnection() = True Then
            Dim myCmd As New FbCommand(sSQL, myConnection)

            Try
                myCmd.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                MessageBox.Show("An error has occured!" & vbCrLf & vbCrLf &
                ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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



    Public Shared Function getDataFromTable(ByVal tableName As String, Optional ByVal columnNames As String = "", Optional ByVal whereClause As String = "", Optional ByVal Disconnect As Boolean = False) As FbDataReader

        Dim conn As New FbConnection
        Dim command As New FbCommand
        Dim ds As FbDataReader = Nothing
        Dim sql As String
        Dim resultList As List(Of String)

        If OpenSQLConnection() = True Then
            Try
                'excecuted the SQL command 
                If columnNames IsNot "" Then
                    sql = "Select " & columnNames & " from USERS"
                Else
                    sql = "Select * from USERS"
                End If

                If whereClause IsNot "" Then
                    sql = sql & " where " & whereClause
                End If
                command.Connection = conn
                command.CommandText = sql

                ds = command.ExecuteReader
                resultList = New List(Of String)
                'loops and add the fields into the table 
                ' While ds.Read()
                'resultList.Add(ds.GetValue(ds.GetOrdinal("INDEX_NAME")))

                'End While

                Return ds
            Catch ex As Exception
                MessageBox.Show("An error has occured!" & vbCrLf & vbCrLf &
                ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try

        End If
        Return ds
    End Function

End Class
