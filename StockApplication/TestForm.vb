Imports FirebirdSql.Data.FirebirdClient
Imports System.Windows.Forms.DataVisualization.Charting

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

        'Dim tmpHourlyStockQuote As HourlyStockQuote
        'tmpHourlyStockQuote = New HourlyStockQuote()
        'tmpHourlyStockQuote.GetAndStoreHourlyData()

        Dim tmpMovingAverageCalculation As MovingAverage
        tmpMovingAverageCalculation = New MovingAverage()
        tmpMovingAverageCalculation.CalculateAndStoreIntraDayMA()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'NSEIndices Data fetch
        'Dim Fdataa As New FbDataAdapter("select * from STOCKHOURLYDATA", "servertype=0;username=sysdba;password=Jan@2017;database=D:\Tarun\StockApp\DB\STOCKAPPDB.fdb;datasource=localhost")

        'DataSet1.Tables.Add("STOCKHOURLYDATA")
        'Me.DataGridView1.DataSource = DataSet1
        'Me.DataGridView1.DataMember = "STOCKHOURLYDATA"

        'Fdataa.Fill(DataSet1, "STOCKHOURLYDATA")

        'Chart1.Series.Add("Price")


        Chart1.ChartAreas.Clear()
        Chart1.Series.Clear()

        'Create a chartarea
        Dim area As New ChartArea("AREA")
        'If you don't set these limits, they will be generated automatically depending on the data.
        area.AxisX.Interval = 1
        'area.AxisX.Minimum = 9
        'area.AxisX.Maximum = 17
        area.AxisY.Minimum = 86.4
        'area.AxisX.LabelStyle.Format = "HH:MM:SS"
        'area.AxisX.IntervalType = DateTimeIntervalType.Minutes


        area.AxisY.Maximum = 89.8
        'Add the chart area to the chart
        Chart1.ChartAreas.Add(area)







        'Create a Series
        Dim StockData As Series = New Series("Price")
        StockData.ChartType = SeriesChartType.Line
        'Assign the series to the area
        StockData.ChartArea = "AREA"
        'Add the series to the chart
        Chart1.Series.Add(StockData)
        Chart1.Series(0).IsXValueIndexed = True
        Dim whereClause, orderClause As String
        Dim ds As FbDataReader = Nothing
        whereClause = "TRADEDDATE = '30-Mar-2017' and companycode = 'IDEA'"
        orderClause = "lastupdatetime"


        ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " lastClosingPrice, lastupdatetime", whereClause, orderClause)
        ' Open database connection
        'Dim Fdataa As New FbDataAdapter("select * from STOCKHOURLYDATA", "servertype=0;username=sysdba;password=Jan@2017;database=D:\Tarun\StockApp\DB\STOCKAPPDB.fdb;datasource=localhost")
        'Dim myConnection As New OleDbConnection(myConnectionString)
        'Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

        'myCommand.Connection.Open()



        'Fdataa.Fill(DataSet1, "STOCKHOURLYDATA")
        ' Create a database reader    
        'Dim myReader As OleDbDataReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)


        While ds.Read()
            Chart1.Series("Price").Points.AddXY(ds("lastupdatetime"), ds("lastClosingPrice"))
        End While

        ' Close the reader and the connection
        ds.Close()
        DBFunctions.CloseSQLConnection()
    End Sub


End Class