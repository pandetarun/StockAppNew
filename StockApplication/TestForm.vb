﻿Imports FirebirdSql.Data.FirebirdClient
Imports FirebirdSql.Data.Isql
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Collections.Generic

Imports System.IO.StreamReader
Imports System.Globalization
Imports System.Data
Imports System.IO
Imports System.Windows.Forms



Public Class TestForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'NSEIndices Testing
        Dim tmpNSEindices As NSEindices
        tmpNSEindices = New NSEindices()
        tmpNSEindices.getIndicesListAndStore()

        'Dim tmpNSEIndicesDetails As NSEIndicesDetails
        'tmpNSEIndicesDetails = New NSEIndicesDetails()
        'tmpNSEIndicesDetails.getIndicesDetailsAndStore()

        'DataLayer testing
        'Dim myDataLayer As DataLayer = New DataLayer()
        'myDataLayer.CreateDatabase()

        'Dim tmpHourlyStockQuote As HourlyStockQuote
        'tmpHourlyStockQuote = New HourlyStockQuote()
        'tmpHourlyStockQuote.GetAndStoreHourlyData()

        'Dim tmpHourlyStockQuote As DailyStockQuote
        'tmpHourlyStockQuote = New DailyStockQuote()
        'tmpHourlyStockQuote.getDailyStockDetailsAndStore()

        'Dim tmpMovingAverageCalculation As MovingAverage
        'tmpMovingAverageCalculation = New MovingAverage()
        'tmpMovingAverageCalculation.CalculateAndStoreIntraDayMA()

        'Dim tmpMovingAverageCalculation As MovingAverage
        'tmpMovingAverageCalculation = New MovingAverage()
        'tmpMovingAverageCalculation.CalculateAndStoreDayMA()

        'Dim tmpMovingAverageCalculation As BollingerBands
        'tmpMovingAverageCalculation = New BollingerBands()
        'tmpMovingAverageCalculation.CalculateAndStoreIntradayBollingerBands()

        'DBFunctions.CreateDatabase()

        'Dim conn As New FbConnection
        'Dim command As New FbBatchExecution
        'Dim ds As FbDataReader = Nothing

        'If DBFunctions.OpenSQLConnection() = True Then
        '    Try
        '        Dim script As String
        '        script = "INSERT INTO TESTTABLE (NAME, SALARY) VALUES  ('Tarun',20 );"
        '        script = script & "INSERT INTO TESTTABLE (NAME, SALARY) VALUES  ('Sangita',23 );"
        '        script = script & "INSERT INTO TESTTABLE (NAME, SALARY) VALUES  ('Sony',12 );"
        '        script = script & "INSERT INTO TESTTABLE (NAME, SALARY) VALUES  ('Neha',34 );"
        '        script = script & "INSERT INTO TESTTABLE (NAME, SALARY) VALUES  ('Roli',3 );"
        '        Dim fbs As FbScript = New FbScript(script)
        '        fbs.Parse()
        '        command = New FbBatchExecution(DBFunctions.myConnection)


        '        command.AppendSqlStatements(fbs)


        '        command.Execute(True)

        '        DBFunctions.CloseSQLConnection()

        '        StockAppLogger.Log("getDataFromTable End", "DBFunctions")

        '    Catch ex As Exception
        '        StockAppLogger.LogError("getDataFromTable Error in getting data for query ", ex, "DBFunctions")

        '    End Try
        'End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub



    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub populateSMAValues()
        Dim ds1 As FbDataReader = Nothing
        Dim tmpMAPeriods, tmpEMAPeriods As List(Of String)
        Dim configuredSMAPeriods, configuredEMAPeriods As String
        Dim counter As Integer = 0


        tmpMAPeriods = Nothing
        tmpEMAPeriods = Nothing

        If ComboBox1.SelectedItem IsNot Nothing Then
            ds1 = DBFunctions.getDataFromTableExt("STOCKWISEPERIODS", "DC", " INTRADAYSMAPERIOD, INTRADAYEMAPERIOD", "stockname = '" & ComboBox1.SelectedItem.ToString() & "'")
            If ds1.Read() Then
                configuredSMAPeriods = ds1.GetValue(ds1.GetOrdinal("INTRADAYSMAPERIOD"))
                configuredEMAPeriods = ds1.GetValue(ds1.GetOrdinal("INTRADAYEMAPERIOD"))
                tmpMAPeriods = New List(Of String)(configuredSMAPeriods.Split(","))
                tmpEMAPeriods = New List(Of String)(configuredEMAPeriods.Split(","))
            End If

            If tmpMAPeriods IsNot Nothing Then
                SMA.Items.Add("Select")
                For counter = 0 To tmpMAPeriods.Count - 1
                    SMA.Items.Add(tmpMAPeriods.Item(counter))
                Next
            End If

            If tmpEMAPeriods IsNot Nothing Then
                EMA.Items.Add("Select")
                For counter = 0 To tmpEMAPeriods.Count - 1
                    EMA.Items.Add(tmpEMAPeriods.Item(counter))
                Next
            End If

        End If

    End Sub

    Private Sub TestForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Select()
        populateSMAValues()
        Dim ds As FbDataReader = Nothing
        Chart1.ChartAreas.Clear()
        Chart1.Series.Clear()
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
        populateSMAValues()
        DrawGraph()
    End Sub

    Private Sub DrawGraph()
        Dim whereClause, orderClause As String
        Dim ds As FbDataReader = Nothing
        Dim MAds As FbDataReader = Nothing
        Dim MAwhereclause As String
        Dim MAOrderby As String

        Chart1.ChartAreas.Clear()
        Chart1.Series.Clear()


        whereClause = "TRADEDDATE = '" & DateTimePicker1.Value.Date & "' and companycode = '" & ComboBox1.SelectedItem.ToString() & "'"
        'orderClause = "lastupdatetime"
        ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " Max(lastClosingPrice) maxprice, Min(lastClosingPrice) minprice", whereClause)
        'Create a chartarea
        Dim area As New ChartArea("AREA")
        'If you don't set these limits, they will be generated automatically depending on the data.
        area.AxisX.Interval = 1
        If ds.Read Then
            area.AxisY.Minimum = ds("minprice") - 1
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
        'Dim SMAData As Series = New Series("SMA")
        'StockData.ChartType = SeriesChartType.Line
        'Assign the series to the area
        StockData.ChartArea = "AREA"
        'SMAData.ChartArea = "AREA"
        'Add the series to the chart
        Chart1.Series.Add(StockData)
        With Chart1.ChartAreas(0)
            '.AxisX.MajorGrid.LineDashStyle = Drawing.Design.
            .AxisX.MajorGrid.LineColor = Drawing.Color.Aqua
            .AxisY.MajorGrid.LineColor = Drawing.Color.Aqua
            .AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet

        End With
        'Chart1.Series.Add(SMAData)
        'Chart1.Series(0).IsXValueIndexed = True
        'Chart1.Series(1).IsXValueIndexed = True
        If SMA.SelectedIndex < 0 Then
            whereClause = "TRADEDDATE = '" & DateTimePicker1.Value.Date & "' and companycode = '" & ComboBox1.SelectedItem.ToString() & "'"
            orderClause = "lastupdatetime"
            ds = DBFunctions.getDataFromTableExt("STOCKHOURLYDATA", "DC", " lastClosingPrice, lastupdatetime", whereClause, orderClause)
            While ds.Read()
                Chart1.Series("Price").Points.AddXY(ds("lastupdatetime"), ds("lastClosingPrice"))
            End While
            ' Close the reader and the connection
            ds.Close()

        ElseIf SMA.SelectedIndex >= 0 Then
            Dim SMAData As Series = New Series("SMA")
            SMAData.ChartType = SeriesChartType.Line
            SMAData.ChartArea = "AREA"
            SMAData.Color = Drawing.Color.Black
            Chart1.Series.Add(SMAData)
            'Chart1.Series(0).IsXValueIndexed = True

            'MAwhereclause = "SHD.TRADEDDATE ='" & DateTimePicker1.Value.Date & "' and SHD.companycode='" & ComboBox1.SelectedItem.ToString() & "' and IMA.stock_name='" & ComboBox1.SelectedItem.ToString() & "' and SHD.companycode = IMA.STOCK_NAME And shd.lastupdatetime = ima.lastupdatetime"
            MAwhereclause = "SHD.TRADEDDATE ='" & DateTimePicker1.Value.Date & "' and SHD.companycode='" & ComboBox1.SelectedItem.ToString() & "'"
            MAOrderby = "SHD.lastupdatetime"
            MAds = DBFunctions.getDataFromTableExt("STOCKHOURLYDATA SHD left join INTRADAYSNEMOVINGAVERAGES IMA on shd.lastupdatetime = ima.lastupdatetime and IMA.TRADEDDATE = SHD.TRADEDDATE and SHD.companycode = IMA.STOCKNAME And ima.period = " & SMA.SelectedItem.ToString, "DC", " SHD.tradeddate, SHD.lastupdatetime lastupdatetime, SHD.LASTCLOSINGPRICE closingprice, IMA.SMA SMA ", MAwhereclause, MAOrderby)
            While MAds.Read()
                Chart1.Series("Price").Points.AddXY(MAds("lastupdatetime"), MAds("closingprice"))
                Chart1.Series("SMA").Points.AddXY(MAds("lastupdatetime"), MAds("SMA"))
            End While
            MAds.Close()
        End If
        DBFunctions.CloseSQLConnectionExt("DC")
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        DrawGraph()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SMA.SelectedIndexChanged
        DrawGraph()
    End Sub
End Class