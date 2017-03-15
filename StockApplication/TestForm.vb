Imports FirebirdSql.Data.FirebirdClient
Public Class TestForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'NSEIndices Testing
        'Dim tmpNSEindices As NSEindices
        'tmpNSEindices = New NSEindices()
        'tmpNSEindices.getIndicesListAndStore()

        'Dim tmpNSEIndicesDetails As NSEIndicesDetails
        'tmpNSEIndicesDetails = New NSEIndicesDetails()
        'tmpNSEIndicesDetails.getIndicesDetailsAndStore()

        'DataLayer testing
        'Dim myDataLayer As DataLayer = New DataLayer()
        'myDataLayer.CreateDatabase()
        Dim tmpHourlyStockQuote As HourlyStockQuote
        tmpHourlyStockQuote = New HourlyStockQuote()
        tmpHourlyStockQuote.GetAndStoreHourlyData()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'NSEIndices Data fetch
        Dim Fdataa As New FbDataAdapter("select * from STOCKHOURLYDATA", "servertype=0;username=sysdba;password=Jan@2017;database=D:\Tarun\StockApp\DB\STOCKAPPDB.fdb;datasource=localhost")

        DataSet1.Tables.Add("STOCKHOURLYDATA")
        Me.DataGridView1.DataSource = DataSet1
        Me.DataGridView1.DataMember = "STOCKHOURLYDATA"

        Fdataa.Fill(DataSet1, "STOCKHOURLYDATA")
    End Sub
End Class