Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class MACD
    Dim MATime, MAStock As String
    Dim MADate As Date
    Dim counter As Integer
    Dim lastTradedPrice As Double

    Public Function CalculateAndStoreMACD() As Boolean
        GetStockListAndCalculateIntraDayMACD()
        Return False
    End Function

    Public Function CalculateAndStoreDayMA() As Boolean
        'GetStockListAndCalculateDailyMA()
        Return False
    End Function

    Private Function GetStockListAndCalculateIntraDayMACD() As Boolean
        Dim tmpStockList As List(Of String) = New List(Of String)
        Dim tmpStockCode As String
        Dim ds As FbDataReader = Nothing

        StockAppLogger.Log("GetStockListAndCalculateIntraDayMACD Start", "MovingAverage")
        Try
            ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING")
            While ds.Read()
                tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
                If Not tmpStockList.Contains(tmpStockCode) Then
                    tmpStockList.Add(tmpStockCode)
                    If IntraDayMACDCalculation(tmpStockCode) Then
                        StockAppLogger.Log("GetStockListAndCalculateIntraDayMACD IntraDaySNEMACalculation successfull for Stock = " & tmpStockCode & " at time = " & MATime, "MovingAverage")
                    Else
                        StockAppLogger.LogInfo("GetStockListAndCalculateIntraDayMACD IntraDaySNEMACalculation failed for Stock = " & tmpStockCode & " at time = " & MATime, "MovingAverage")
                    End If
                End If
            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("GetStockListAndCalculateIntraDayMACD Error Occurred in getting stocklist from DB = ", exc, "MovingAverage")
            Return False
        End Try
        StockAppLogger.Log("GetStockListAndCalculateIntraDayMACD End", "MovingAverage")
        Return True
    End Function

    Private Function IntraDayMACDCalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim orderClause, whereClause, configuredMACDPeriods As String
        Dim closingPrice As Double
        Dim laststoredEMA, tmpSimpleMA As Double
        Dim tmpMACDPeriods As List(Of String)

        StockAppLogger.Log("IntraDayMACDCalculation Start", "MovingAverage")

        tmpMACDPeriods = Nothing
        lastTradedPrice = 0
        counter = 0
        'simpleMA = 0
        tmpSimpleMA = 0
        laststoredEMA = 0
        MAStock = tmpStockCode
        Try
            whereClause = "TRADEDDATE='" & Today & "' and STOCKNAME = '" & tmpStockCode & "' and period"
            orderClause = "lastupdatetime desc"
            MADate = Today
            ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " lastClosingPrice, lastupdatetime", whereClause, orderClause)
            'Get period details for the stock
            ds1 = DBFunctions.getDataFromTable("STOCKWISEPERIODS", " MACD", "stockname = '" & tmpStockCode & "'")
            If ds1.Read() Then
                configuredMACDPeriods = ds1.GetValue(ds1.GetOrdinal("MACD"))
                tmpMACDPeriods = New List(Of String)(configuredMACDPeriods.Split(","))
            End If
            '    While ds.Read()
            '        counter = counter + 1
            '        If counter = 1 Then
            '            MATime = ds.GetValue(ds.GetOrdinal("lastupdatetime"))
            '            closingPrice = ds.GetValue(ds.GetOrdinal("lastClosingPrice"))
            '            lastTradedPrice = closingPrice
            '        End If
            '        tmpSimpleMA = tmpSimpleMA + Double.Parse(ds.GetValue(ds.GetOrdinal("lastClosingPrice")))
            '        If tmpMACDPeriods.Contains(counter) Then
            '            simpleMA = tmpSimpleMA / counter
            '            CalculateEMAForPeriod(counter, tmpStockCode, closingPrice)
            '            If InsertIntraDaySNEMAtoDB(counter) Then
            '                StockAppLogger.Log("IntraDaySNEMACalculation period " & counter & " inserted in DB for stock = " & tmpStockCode, "MovingAverage")
            '            Else
            '                StockAppLogger.LogInfo("IntraDaySNEMACalculation period " & counter & " not inserted in DB  for stock = " & tmpStockCode, "MovingAverage")
            '            End If
            '            If tmpMAPeriods.IndexOf(counter) = tmpMAPeriods.Count Then
            '                Exit While
            '            End If
            '        End If
            '    End While
            '    If counter = 1 Then
            '        simpleMA = tmpSimpleMA
            '        If InsertIntraDaySNEMAtoDB(counter) Then
            '            StockAppLogger.Log("IntraDaySNEMACalculation period " & counter & " inserted in DB for stock = " & tmpStockCode, "MovingAverage")
            '        Else
            '            StockAppLogger.LogInfo("IntraDaySNEMACalculation period " & counter & " not inserted in DB  for stock = " & tmpStockCode, "MovingAverage")
            '        End If
            '    End If
        Catch exc As Exception
            StockAppLogger.LogError("IntraDaySNEMACalculation Error Occurred in calculating intraday moving average = ", exc, "MovingAverage")
            Return False
        End Try
        StockAppLogger.Log("IntraDaySNEMACalculation End", "MovingAverage")
        Return True
    End Function

End Class
