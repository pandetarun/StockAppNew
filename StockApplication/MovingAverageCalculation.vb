Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class MovingAverageCalculation
    Dim threeSampleMinSMA, fiveSampleMinSMA, tenSampleMinSMA, fifteenSampleSMA, twentySampleMinSMA, FiftySampleMinSMA, simpleMA As Double
    Dim threeSampleMinEMA, fiveSampleMinEMA, tenSampleMinEMA, fifteenSampleEMA, twentySampleMinEMA, FiftySampleMinEMA As Double
    Dim MATime, MAStock As String
    Dim MADate As Date
    Dim counter As Integer

    Public Function CalculateAndStoreIntraDayMA() As Boolean
        Dim tmpStockList As List(Of String) = New List(Of String)

        tmpStockList = GetStockList()
        For Each tmpStock In tmpStockList

        Next
        Return False
    End Function

    Private Function GetStockList() As List(Of String)
        Dim tmpStockList As List(Of String) = New List(Of String)
        Dim tmpStockCode As String
        Dim ds As FbDataReader = Nothing

        StockAppLogger.Log("GetStockList Start", "MovingAverageCalculation")
        Try
            ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING")
            While ds.Read()
                tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
                If Not tmpStockList.Contains(tmpStockCode) Then
                    tmpStockList.Add(tmpStockCode)
                    If HourlyMACalculation(tmpStockCode) Then

                    Else
                        StockAppLogger.Log("GetStockList HourlyMACalcultaion failed for Stock = " & tmpStockCode & " at time = " & MATime, "MovingAverageCalculation")
                    End If
                End If
            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting stocklist from DB = ", exc, "MovingAverageCalculation")
        End Try
        StockAppLogger.Log("GetStockList End", "MovingAverageCalculation")
        Return tmpStockList
    End Function

    Private Function HourlyMACalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim whereClause, whereClause1 As String
        Dim orderClause, orderClause1 As String
        Dim closingPrice As Double
        Dim laststoredEMA As Double

        counter = 0
        simpleMA = 0
        threeSampleMinSMA = 0
        fiveSampleMinSMA = 0
        tenSampleMinSMA = 0
        fifteenSampleSMA = 0
        twentySampleMinSMA = 0
        FiftySampleMinSMA = 0

        threeSampleMinEMA = 0
        fiveSampleMinEMA = 0
        tenSampleMinEMA = 0
        fifteenSampleEMA = 0
        twentySampleMinEMA = 0
        FiftySampleMinEMA = 0

        StockAppLogger.Log("HourlyMACalculation Start", "MovingAverageCalculation")
        Try
            whereClause1 = "TRADEDDATE='" & Today & "' and STOCK_NAME = '" & tmpStockCode & "'"
            orderClause1 = "LASTUPDATETIME desc"
            whereClause = "TRADEDDATE='" & Today & "' and companycode = '" & tmpStockCode & "'"
            orderClause = "lastupdatetime desc"


            'orderClause1 = "lastupdatetime"
            ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " lastClosingPrice, lastupdatetime", whereClause, orderClause)
            ds1 = DBFunctions.getDataFromTable("INTRADAYMOVINGAVERAGES", " * ", whereClause1, orderClause1)
            MATime = ds.GetValue(ds.GetOrdinal("lastupdatetime"))
            closingPrice = ds.GetValue(ds.GetOrdinal("lastClosingPrice"))
            ds1.Read()
            While ds.Read()
                counter = counter + 1
                simpleMA = simpleMA + Double.Parse(ds.GetValue(ds.GetOrdinal("lastClosingPrice")))
                If counter = 3 Then
                    threeSampleMinSMA = simpleMA
                    laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("3EMA")))
                    threeSampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                ElseIf counter = 5 Then
                    fiveSampleMinSMA = simpleMA / 5
                    laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("5EMA")))
                    fiveSampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                ElseIf counter = 10 Then
                    tenSampleMinSMA = simpleMA / 10
                    laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("10EMA")))
                    tenSampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                ElseIf counter = 15 Then
                    fifteenSampleSMA = simpleMA / 15
                    laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("15EMA")))
                    fifteenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                ElseIf counter = 20 Then
                    twentySampleMinSMA = simpleMA / 20
                    laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("20EMA")))
                    twentySampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                ElseIf counter = 50 Then
                    FiftySampleMinSMA = simpleMA / 50
                    laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("50EMA")))
                    FiftySampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Exit While
                End If
            End While
            If counter = 1 Then
                threeSampleMinSMA = simpleMA
                fiveSampleMinSMA = simpleMA
                tenSampleMinSMA = simpleMA
                fifteenSampleSMA = simpleMA
                twentySampleMinSMA = simpleMA
                FiftySampleMinSMA = simpleMA
                threeSampleMinEMA = simpleMA
                fiveSampleMinEMA = simpleMA
                tenSampleMinEMA = simpleMA
                fifteenSampleEMA = simpleMA
                twentySampleMinEMA = simpleMA
                FiftySampleMinEMA = simpleMA
            End If
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting hourly stock data from DB = ", exc, "MovingAverageCalculation")
            Return False
        End Try
        StockAppLogger.Log("HourlyMACalculation End", "MovingAverageCalculation")
        Return True
    End Function

    Private Function InsertMAtoDB() As Boolean

        Dim insertStatement As String
        Dim insetValues As String
        Dim sqlStatement As String

        StockAppLogger.Log("InsertMAtoDB Start", "MovingAverageCalculation")
        insertStatement = "INSERT INTO INTRADAYMOVINGAVERAGES (DATE, TIME, STOCK_NAME"
        insetValues = "VALUES ('" & MADate & "','" & MATime & "', '" & MAStock & "'."
        If threeSampleMinSMA <> 0 Then
            insertStatement = insertStatement & ", 3MA"
            insetValues = insetValues & threeSampleMinSMA
        End If
        If fiveSampleMinSMA <> 0 Then
            insertStatement = insertStatement & ", 5MA"
            insetValues = insetValues & fiveSampleMinSMA
        End If
        If tenSampleMinSMA <> 0 Then
            insertStatement = insertStatement & ", 10MA"
            insetValues = insetValues & tenSampleMinSMA
        End If
        If fifteenSampleSMA <> 0 Then
            insertStatement = insertStatement & ", 15MA"
            insetValues = insetValues & fifteenSampleSMA
        End If
        If twentySampleMinSMA <> 0 Then
            insertStatement = insertStatement & ", 20MA"
            insetValues = insetValues & twentySampleMinSMA
        End If
        If FiftySampleMinSMA <> 0 Then
            insertStatement = insertStatement & ", 50MA) "
            insetValues = insetValues & FiftySampleMinSMA
        End If

        If threeSampleMinEMA <> 0 Then
            insertStatement = insertStatement & ", 3EMA"
            insetValues = insetValues & threeSampleMinEMA
        End If
        If fiveSampleMinEMA <> 0 Then
            insertStatement = insertStatement & ", 5EMA"
            insetValues = insetValues & fiveSampleMinEMA
        End If
        If tenSampleMinEMA <> 0 Then
            insertStatement = insertStatement & ", 10EMA"
            insetValues = insetValues & tenSampleMinEMA
        End If
        If fifteenSampleEMA <> 0 Then
            insertStatement = insertStatement & ", 15EMA"
            insetValues = insetValues & fifteenSampleEMA
        End If
        If twentySampleMinEMA <> 0 Then
            insertStatement = insertStatement & ", 20EMA"
            insetValues = insetValues & twentySampleMinEMA
        End If
        If FiftySampleMinEMA <> 0 Then
            insertStatement = insertStatement & ", 50EMA) "
            insetValues = insetValues & FiftySampleMinEMA & ");"
        End If

        sqlStatement = insertStatement & insetValues
        DBFunctions.ExecuteSQLStmt(sqlStatement)
        StockAppLogger.Log("InsertMAtoDB End", "MovingAverageCalculation")
        Return True
    End Function
End Class
