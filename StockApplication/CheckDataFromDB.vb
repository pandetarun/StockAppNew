Imports FirebirdSql.Data.FirebirdClient
Imports FirebirdSql.Data.Isql
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Collections.Generic

Imports System.IO.StreamReader
Imports System.Globalization
Imports System.Data
Imports System.IO
Imports System.Windows.Forms

Public Class CheckDataFromDB
    Private Sub CheckDataFromDB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Select()

        Dim ds As FbDataReader = Nothing

        'whereClause = "TRADEDDATE = '30-Mar-2017' and companycode = 'IDEA'"

        ds = DBFunctions.getDataFromTableExt("NSE_INDICES_TO_STOCK_MAPPING", "DC", " distinct STOCK_NAME")

        If ds.HasRows Then
            Do While ds.Read()
                ComboBox1.Items.Add(ds("STOCK_NAME"))
            Loop
        End If
        ds.Close()
        DBFunctions.CloseSQLConnectionExt("DC")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        populateTableValues()

    End Sub

    Private Sub populateTableValues()
        Dim ds As FbDataReader = Nothing
        Dim ssql As String

        ssql = "select * from STOCKHOURLYDATA where companycode = '" & ComboBox1.SelectedItem.ToString() & "'"
        ds = DBFunctions.ExecuteSQLStmtandReturnResultExt(ssql, "DC")

        Dim dt = New DataTable()
        dt.Load(ds)
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.DataSource = dt
        DataGridView1.Refresh()
        ' DataGridView1.DataSource = DataSet1.Tables(ComboBox1.SelectedItem.ToString())

    End Sub
End Class