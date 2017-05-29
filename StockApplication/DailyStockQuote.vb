Imports System.Collections.Generic
Imports System
Imports System.IO

Public Class DailyStockQuote

    Dim symbol As String
    Dim lastPrice As Double
    Dim changeInPercentage As Double
    Dim volume As Double
    Dim highPrice As Double
    Dim lowPrice As Double
    Dim yearlyHighPrice As Double
    Dim yearlyLowPrice As Double

    Dim updateDate As Date
    Dim updateTime As String
    Dim pathOfFileofQuots As String = My.Settings.ApplicationFileLocation & "\Quotes.txt"
    Dim dailyStockDetailsList As List(Of DailyStockQuote) = New List(Of DailyStockQuote)
    Dim stockList As List(Of String) = New List(Of String)

    Public Function getDailyStockDetailsAndStore() As Boolean

        StockAppLogger.Log("getDailyStockDetailsAndStore Start", "DailyStockQuote")
        Try
            readQuotesFromFile()
            StoreDailyStockDetail()
            DBFunctions.CloseSQLConnectionExt("DC")
        Catch ex As Exception
            StockAppLogger.LogError("getDailyStockDetailsAndStore Error ", ex, "DailyStockQuote")
            DBFunctions.CloseSQLConnectionExt("DC")
            Return False
        End Try
        StockAppLogger.Log("getDailyStockDetailsAndStore End", "DailyStockQuote")
        Return True
    End Function

    Private Sub readQuotesFromFile()
        Dim reader = File.OpenText(pathOfFileofQuots)
        Dim line As String = Nothing
        Dim lineArray As String()
        Dim dailyStockQuotObj As DailyStockQuote
        StockAppLogger.Log("readQuotesFromFile Start", "DailyStockQuote")
        While (reader.Peek() <> -1)
            line = reader.ReadLine()
            lineArray = line.Split("\t")
            dailyStockQuotObj = New DailyStockQuote()
            dailyStockQuotObj.symbol = lineArray(0)
            dailyStockQuotObj.lastPrice = lineArray(1)
            dailyStockQuotObj.changeInPercentage = lineArray(2)
            dailyStockQuotObj.volume = lineArray(3)
            dailyStockQuotObj.highPrice = lineArray(4)
            dailyStockQuotObj.lowPrice = lineArray(5)
            dailyStockQuotObj.yearlyHighPrice = lineArray(6)
            dailyStockQuotObj.yearlyLowPrice = lineArray(7)
            dailyStockDetailsList.Add(dailyStockQuotObj)
        End While
        StockAppLogger.Log("readQuotesFromFile End", "DailyStockQuote")
    End Sub

    Private Sub StoreDailyStockDetail()
        Dim insertStatement As String
        Dim insertValues As String

        StockAppLogger.Log("StoreDailyStockDetail Start", "DailyStockQuote")
        Try
            For Each tmpDailyStockDetails In dailyStockDetailsList
                StockAppLogger.Log("StoreDailyStockDetail storing daily data for stock = " & tmpDailyStockDetails.symbol, "DailyStockQuote")
                insertStatement = "INSERT INTO DAILYSTOCKDATA (STOCKNAME, OPENPRICE, HIGHPRICE, LOWPRICE, LAST_TRADED_PRICE, CHANGE, CHANGE_PERCENTAGE, VOLUME, TURNOVER_IN_CRS, YEARLY_HIGH, YEARLY_LOW, YERLY_PERCENTAGE_CHANGE, MONTHLY_PERCENTAGE_CHANGE, TRADEDDATE)"
                insertValues = " VALUES ('"
                insertValues = insertValues & tmpDailyStockDetails.symbol & "',"
                insertValues = insertValues & tmpDailyStockDetails.openPrice & ","
                insertValues = insertValues & tmpDailyStockDetails.highPrice & ","
                insertValues = insertValues & tmpDailyStockDetails.lowPrice & ","
                insertValues = insertValues & tmpDailyStockDetails.lastTradedPrice & ","
                insertValues = insertValues & tmpDailyStockDetails.changeinPrice & ","
                insertValues = insertValues & tmpDailyStockDetails.changeInPercentage & ","
                insertValues = insertValues & tmpDailyStockDetails.volume & ","
                insertValues = insertValues & tmpDailyStockDetails.turnover & ","
                insertValues = insertValues & tmpDailyStockDetails.yearlyHighPrice & ","
                insertValues = insertValues & tmpDailyStockDetails.yearlyLowPrice & ","
                insertValues = insertValues & tmpDailyStockDetails.yearlyPercentageChange & ","
                insertValues = insertValues & tmpDailyStockDetails.monthlyPercentageChange & ","
                insertValues = insertValues & "'" & tmpDailyStockDetails.updateDate & "');"
                DBFunctions.ExecuteSQLStmtExt(insertStatement & insertValues, "DC")
            Next


        Catch ex As Exception
            StockAppLogger.LogError("StoreDailyStockDetail Error in storing daily record", ex, "DailyStockQuote")
        End Try
        StockAppLogger.Log("StoreDailyStockDetail End", "DailyStockQuote")
    End Sub

End Class
