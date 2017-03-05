'Imports FirebirdSql.Data.Firebird
Imports FirebirdSql.Data.FirebirdClient
Imports System.IO.StreamReader
Imports System.Globalization
Imports System.Data
Imports System.IO
Imports System.Windows

Public Class DataLayer
    Public myCellCollection As New Collection
    Public myConnection As New FbConnection
    Public myDataAdapter() As FbDataAdapter
    Public myDataSet As New DataSet
    Public myWorkRow As DataRow

    Public ServerType As FbServerType = FbServerType.Embedded   '!
    Public Database As String = "D:\Tarun\StockApp\StockAppDB.fdb"                   '!
    Public DataSource As String = "FireBird"                            '!
    Public Password As String = "Jan@2017"                              '!
    Public UserID As String = "SYSDBA"                                '!

    Public Function OpenSQLConnection() As Boolean
        Try
            If myConnection.State = ConnectionState.Closed Then
                Dim cs As New FbConnectionStringBuilder()

                ' If Not ServerType = FbServerType.Embedded Then
                cs.DataSource = DataSource
                cs.Password = Password
                cs.UserID = UserID
                cs.Port = 3050
                'End If

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

            If Not ServerType = FbServerType.Embedded Then
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

    Public Sub CloseSQLConnection()

        If myConnection.State = ConnectionState.Open Then
            myConnection.Close()
        End If

    End Sub

    Public Sub DisposeDAL()

        If myDataAdapter IsNot Nothing Then
            For a As Short = 0 To UBound(myDataAdapter)
                myDataAdapter(a).Dispose()
                myDataAdapter(a) = Nothing
            Next
            myDataAdapter = Nothing
        End If

        If myDataSet IsNot Nothing Then
            myDataSet.Dispose()
            myDataSet = Nothing
        End If

        If myConnection IsNot Nothing Then
            myConnection = Nothing
        End If

        myCellCollection = Nothing

    End Sub

    Public Function ExecuteSQLStmt(ByVal sSQL As String, Optional ByVal Disconnect As Boolean = False) As Boolean

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

    Public Function StreamDataIntoControl(ByVal sSQL As String, ByVal Col As Short, ByVal myControl As Object, Optional ByVal Disconnect As Boolean = False) As Boolean
        ' This function is compatible with ListBox and ComboBox only.
        '
        If OpenSQLConnection() = True Then
            Dim myCmd As New FbCommand(sSQL, myConnection)

            Try
                Dim myReader As FbDataReader
                myReader = myCmd.ExecuteReader()

                myControl.Items.Clear()
                Do While (myReader.Read())
                    myControl.Items.Add(myReader(Col).ToString())
                Loop

                If myReader IsNot Nothing Then
                    myReader.Close()
                    myReader = Nothing
                End If
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

    Public Function RetrieveTableRecords(ByVal TableIndex As Byte, ByVal sSQL As String, Optional ByVal myControl As Object = Nothing, Optional ByVal Disconnect As Boolean = False) As Boolean
        'An optional myControl can be DataGridView control
        '
        If OpenSQLConnection() = True Then
            Try
                If TableIndex + 1 > myDataSet.Tables.Count Then
                    ReDim Preserve myDataAdapter(TableIndex)
                    myDataSet.Tables.Add()
                End If

                myDataSet.Tables(TableIndex).Clear()
                myDataAdapter(TableIndex) = New FbDataAdapter(sSQL, myConnection)
                myDataAdapter(TableIndex).Fill(myDataSet.Tables(TableIndex))
                If Disconnect = True Then CloseSQLConnection()

                If myControl IsNot Nothing Then
                    myControl.DataSource = myDataSet.Tables(TableIndex)
                End If
                Return True
            Catch ex As Exception
                MessageBox.Show("An error has occured!" & vbCrLf & vbCrLf &
                ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        Else
            Return False
        End If

    End Function

    Public Function UpdateTableRecords(ByVal TableIndex As Byte, ByVal State As DataRowState) As Boolean

        Dim myDataRowsCommandBuilder As New FbCommandBuilder(myDataAdapter(TableIndex))
        Dim ChildRecords As DataTable = myDataSet.Tables(TableIndex).GetChanges(State)

        Try
            Select Case State
                Case DataRowState.Added
                    myDataAdapter(TableIndex).InsertCommand = myDataRowsCommandBuilder.GetInsertCommand
                Case DataRowState.Deleted
                    myDataAdapter(TableIndex).DeleteCommand = myDataRowsCommandBuilder.GetDeleteCommand
                Case DataRowState.Modified
                    myDataAdapter(TableIndex).UpdateCommand = myDataRowsCommandBuilder.GetUpdateCommand
            End Select

            myDataAdapter(TableIndex).Update(ChildRecords)
            myDataSet.Tables(TableIndex).AcceptChanges()
            Return True
        Catch ex As Exception
            MessageBox.Show("An error has occured!" & vbCrLf & vbCrLf &
            ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            If myDataRowsCommandBuilder IsNot Nothing Then
                myDataRowsCommandBuilder.Dispose()
                myDataRowsCommandBuilder = Nothing
            End If
            If ChildRecords IsNot Nothing Then
                ChildRecords.Dispose()
                ChildRecords = Nothing
            End If
        End Try

    End Function

    Public Function AddDataRow(ByVal TableIndex As Byte) As Boolean

        Try
            Dim Cell As Short
            With myDataSet.Tables(TableIndex)

                myWorkRow = .NewRow
                For Cell = 0 To myCellCollection.Count - 1
                    myWorkRow(Cell) = myCellCollection.Item(Cell + 1)
                Next

                myCellCollection.Clear()
                .Rows.Add(myWorkRow)
                Return True
            End With
        Catch ex As Exception
            MessageBox.Show("An error has occured!" & vbCrLf & vbCrLf &
            ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    Public Function DeleteDataRow(ByVal TableIndex As Byte, ByVal Row As Short) As Boolean

        Try
            myDataSet.Tables(TableIndex).Rows(Row).Delete()
            Return True
        Catch ex As Exception
            MessageBox.Show("An error has occured!" & vbCrLf & vbCrLf &
            ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    Public Function GetCellData(ByVal TableIndex As Byte, ByVal Row As Short, ByVal Col As Short) As Object

        Try
            Return myDataSet.Tables(TableIndex).Rows(Row).Item(Col)
        Catch ex As Exception
            MessageBox.Show("An error has occured!" & vbCrLf & vbCrLf &
            ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return "error"
        End Try

    End Function

    Public Function PutCellData(ByVal TableIndex As Byte, ByVal Row As Short, ByVal Col As Short, ByVal NewData As Object) As Boolean

        Try
            myWorkRow = myDataSet.Tables(TableIndex).Rows(Row)
            myWorkRow(Col) = NewData
            Return True
        Catch ex As Exception
            MessageBox.Show("An error has occured!" & vbCrLf & vbCrLf &
            ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    Public Function ImportData(ByVal sSql As String, ByVal Filename As String, Optional ByVal Delimiter As String = ";", Optional ByVal Disconnect As Boolean = False) As Boolean
        'This function imports diskfile into a database
        'Usage: ImportData("INSERT INTO Database (Field1, Field2)", "c:\data.txt", [;])
        '
        If File.Exists(Filename) = True Then
            If OpenSQLConnection() = True Then
                Dim objReader As New StreamReader(Filename)
                Dim myCmd As New FbCommand()
                Dim ValuesPart As String

                Try
                    myCmd.Connection = myConnection
                    Do While objReader.Peek() <> -1
                        ValuesPart = " VALUES (" & objReader.ReadLine() & ")"
                        myCmd.CommandText = sSql & ValuesPart
                        myCmd.ExecuteNonQuery()
                    Loop
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
                    If objReader IsNot Nothing Then
                        objReader.Close()
                        objReader.Dispose()
                        objReader = Nothing
                    End If
                End Try
            Else
                Return False
            End If
        Else
            MessageBox.Show("File does not exist!", "Error",
            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

    End Function
End Class
