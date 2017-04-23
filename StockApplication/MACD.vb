Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class MACD
    Dim MATime, MAStock As String
    Dim MADate As Date
    Dim counter As Integer
    Dim lastTradedPrice As Double
    Dim signalLine, MACD As Double

    Public Function CalculateAndStoreIntraDayMACD() As Boolean
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
            'DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("GetStockListAndCalculateIntraDayMACD Error Occurred in getting stocklist from DB = ", exc, "MovingAverage")
            Return False
        End Try
        StockAppLogger.Log("GetStockListAndCalculateIntraDayMACD End", "MovingAverage")
        Return True
    End Function

    Public Function IntraDayMACDCalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim orderClause, whereClause, configuredMACDPeriods As String
        Dim fastEMA, slowEMA As Double
        Dim tmpMACDPeriods As List(Of String)

        StockAppLogger.Log("IntraDayMACDCalculation Start", "MACD")
        MACD = 0
        fastEMA = 0
        slowEMA = 0
        tmpMACDPeriods = Nothing
        lastTradedPrice = 0
        MAStock = tmpStockCode
        Try
            orderClause = "lastupdatetime desc"
            MADate = Today
            'Get period details for the stock
            ds1 = DBFunctions.getDataFromTable("STOCKWISEPERIODS", " MACD", "stockname = '" & tmpStockCode & "'")
            If ds1.Read() Then
                configuredMACDPeriods = ds1.GetValue(ds1.GetOrdinal("MACD"))
                tmpMACDPeriods = New List(Of String)(configuredMACDPeriods.Split(","))
            End If
            ds1.Close()
            If tmpMACDPeriods IsNot Nothing Then
                whereClause = "TRADEDDATE='" & Today & "' and STOCKNAME = '" & tmpStockCode & "' and period = " & tmpMACDPeriods.Item(0)
                ds = DBFunctions.getDataFromTable("INTRADAYSNEMOVINGAVERAGES", " EMA ", whereClause, orderClause)
                If ds.Read() Then
                    signalLine = ds.GetValue(ds.GetOrdinal("EMA"))
                End If
                ds.Close()
                whereClause = "TRADEDDATE='" & Today & "' and STOCKNAME = '" & tmpStockCode & "' and period = " & tmpMACDPeriods.Item(1)
                ds = DBFunctions.getDataFromTable("INTRADAYSNEMOVINGAVERAGES", " EMA ", whereClause, orderClause)
                If ds.Read() Then
                    fastEMA = ds.GetValue(ds.GetOrdinal("EMA"))
                End If
                ds.Close()
                whereClause = "TRADEDDATE='" & Today & "' and STOCKNAME = '" & tmpStockCode & "' and period = " & tmpMACDPeriods.Item(2)
                ds = DBFunctions.getDataFromTable("INTRADAYSNEMOVINGAVERAGES", " EMA ", whereClause, orderClause)
                If ds.Read() Then
                    slowEMA = ds.GetValue(ds.GetOrdinal("EMA"))
                End If
                ds.Close()
                If fastEMA <> 0 And slowEMA <> 0 Then
                    MACD = fastEMA - slowEMA
                    InsertIntraDayMACDtoDB()
                End If

            End If
        Catch exc As Exception
            StockAppLogger.LogError("IntraDayMACDCalculation Error Occurred in calculating MACD = ", exc, "MACD")
            Return False
        End Try
        StockAppLogger.Log("IntraDayMACDCalculation End", "MACD")
        Return True
    End Function

    Private Function InsertIntraDayMACDtoDB() As Boolean

        Dim insertStatement As String
        Dim insertValues As String
        Dim sqlStatement As String
        Dim fireQuery As Boolean = False

        StockAppLogger.Log("InsertIntraDayMACDtoDB Start", "MACD")
        insertStatement = "INSERT INTO INTRADAYMACD (TRADEDDATE, LASTUPDATETIME, STOCKNAME, SIGNAL, MACD"
        insertValues = "VALUES ('" & MADate & "','" & MATime & "', '" & MAStock & "'," & signalLine & ", " & MACD

        insertStatement = insertStatement & ") "
        insertValues = insertValues & ");"
        sqlStatement = insertStatement & insertValues

        DBFunctions.ExecuteSQLStmt(sqlStatement)

        StockAppLogger.Log("InsertIntraDayMACDtoDB End", "MACD")
        Return True
    End Function

End Class
