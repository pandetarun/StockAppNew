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

        Dim tmpHourlyStockQuote As DailyStockQuote
        tmpHourlyStockQuote = New DailyStockQuote()
        tmpHourlyStockQuote.getDailyStockDetailsAndStore()

        'Dim tmpMovingAverageCalculation As MovingAverage
        'tmpMovingAverageCalculation = New MovingAverage()
        'tmpMovingAverageCalculation.CalculateAndStoreIntraDayMA()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub



    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim whereClause, orderClause As String
        Dim ds As FbDataReader = Nothing
        Chart1.ChartAreas.Clear()
        Chart1.Series.Clear()
        whereClause = "TRADEDDATE = '30-Mar-2017' and companycode = '" & ListBox1.SelectedItem.ToString() & "'"

        ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " Max(lastClosingPrice) maxprice, Min(lastClosingPrice) minprice", whereClause)
        'Create a chartarea
        Dim area As New ChartArea("AREA")
        'If you don't set these limits, they will be generated automatically depending on the data.
        area.AxisX.Interval = 1
        If ds.Read Then
            area.AxisY.Minimum = ds("minprice") + -1
            area.AxisY.Maximum = ds("maxprice") + 1
            area.AxisY.Interval = (ds("maxprice") - ds("minprice")) / 10
        Else
            area.AxisY.Minimum = 80
            area.AxisY.Maximum = 100
        End If
        ds.Close()
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

        whereClause = "TRADEDDATE = '30-Mar-2017' and companycode = '" & ListBox1.SelectedItem.ToString() & "'"
        orderClause = "lastupdatetime"
        ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " lastClosingPrice, lastupdatetime", whereClause, orderClause)
        While ds.Read()
            Chart1.Series("Price").Points.AddXY(ds("lastupdatetime"), ds("lastClosingPrice"))
        End While
        ' Close the reader and the connection
        ds.Close()
        DBFunctions.CloseSQLConnection()
    End Sub

    Private Sub TestForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.ClearSelected()

        Dim ds As FbDataReader = Nothing
        Chart1.ChartAreas.Clear()
        Chart1.Series.Clear()
        'whereClause = "TRADEDDATE = '30-Mar-2017' and companycode = 'IDEA'"

        ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING", " distinct STOCK_NAME")

        If ds.HasRows Then
            Do While ds.Read()
                ListBox1.Items.Add(ds("STOCK_NAME"))
            Loop
        End If
        ds.Close()
        DBFunctions.CloseSQLConnection()
    End Sub
End Class