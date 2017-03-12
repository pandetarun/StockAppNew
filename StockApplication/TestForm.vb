Imports FirebirdSql.Data.FirebirdClient
Public Class TestForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tmpNSEindices As NSEindices
        tmpNSEindices = New NSEindices()
        tmpNSEindices.getIndicesListAndStore()




        'Dim myDataLayer As DataLayer = New DataLayer()
        'myDataLayer.CreateDatabase()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Fdataa As New FbDataAdapter("select * from NSEINDICES", "servertype=0;username=sysdba;password=Jan@2017;database=D:\Tarun\StockApp\DB\STOCKAPPDB.fdb;datasource=localhost")

        DataSet1.Tables.Add("NSEINDICES")
        Me.DataGridView1.DataSource = DataSet1
        Me.DataGridView1.DataMember = "NSEINDICES"

        Fdataa.Fill(DataSet1, "NSEINDICES")
    End Sub
End Class