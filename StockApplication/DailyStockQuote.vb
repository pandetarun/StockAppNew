Imports System.Collections.Generic
Imports System
Imports System.IO

Public Class DailyStockQuote

    Dim symbol As String
    Dim openPrice As Double
    Dim highPrice As Double
    Dim lowPrice As Double
    Dim lastTradedPrice As Double
    Dim changeinPrice As Double
    Dim changeInPercentage As Double
    Dim volume As Double
    Dim turnover As Double
    Dim yearlyHighPrice As Double
    Dim yearlyLowPrice As Double
    Dim yearlyPercentageChange As Double
    Dim monthlyPercentageChange As Double
    Dim updateDate As Date
    Dim updateTime As String

    Dim pathOfFileofURLs As String = My.Settings.ApplicationFileLocation & "\URLs.txt"
    Dim dailyStockDetailsList As List(Of DailyStockQuote) = New List(Of DailyStockQuote)
    Dim stockList As List(Of String) = New List(Of String)

    Public Function getDailyStockDetailsAndStore() As Boolean
        Dim NSEindicesUrlsList As List(Of String)
        StockAppLogger.Log("getDailyStockDetailsAndStore Start", "DailyStockQuote")
        NSEindicesUrlsList = GetNSEUrlForAllIndices()
        getAllStockDetails(NSEindicesUrlsList)

        StockAppLogger.Log("getDailyStockDetailsAndStore End", "DailyStockQuote")
        Return True
    End Function

    Private Function GetNSEUrlForAllIndices() As List(Of String)
        Dim NSEindicesUrlsList As List(Of String) = New List(Of String)
        Dim reader = File.OpenText(pathOfFileofURLs)
        Dim line As String = Nothing
        StockAppLogger.Log("GetNSEUrlForAllIndices Start", "DailyStockQuote")
        While (reader.Peek() <> -1)
            line = reader.ReadLine()
            NSEindicesUrlsList.Add(line)
        End While
        StockAppLogger.Log("GetNSEUrlForAllIndices End", "DailyStockQuote")
        Return NSEindicesUrlsList
    End Function

    Private Function getAllStockDetails(ByVal NSEindicesUrlsList As List(Of String)) As Boolean

        StockAppLogger.Log("getAllStockDetails Start", "DailyStockQuote")
        For Each NSEIndicesURL In NSEindicesUrlsList
            getDailyStockDetails(NSEIndicesURL)
        Next
        StoreDailyStockDetail()
        StockAppLogger.Log("getAllStockDetails End", "DailyStockQuote")
        Return True
    End Function

    Private Sub getDailyStockDetails(ByVal NSEIndicesURL As String)
        Dim rawIndicesData As String
        Dim tmpDailyStockQuote As DailyStockQuote
        Dim tmpRawStockDailyData As String
        Dim indexOfVar As Integer
        Dim tmpString As String
        Dim lastUpdatedDate As Date
        Dim lastupdatedTime As String

        StockAppLogger.Log("getDailyStockDetails Start", "DailyStockQuote")
        Try { 
            rawIndicesData = Helper.GetDataFromUrl(NSEIndicesURL)
            tmpRawStockDailyData = rawIndicesData
            lastUpdatedDate = Today
            indexOfVar = tmpRawStockDailyData.IndexOf("time")
            tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 6)
            tmpString = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))
            lastupdatedTime = tmpString.Substring(tmpString.Length - 8)

            While tmpRawStockDailyData.IndexOf("symbol") >= 1
                If Not stockList.Contains(tmpRawStockDailyData.IndexOf("symbol")) Then
                    tmpDailyStockQuote = New DailyStockQuote()
                    tmpDailyStockQuote.updateDate = lastUpdatedDate
                    tmpDailyStockQuote.updateTime = lastupdatedTime
                    indexOfVar = tmpRawStockDailyData.IndexOf("symbol")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 9)
                    tmpDailyStockQuote.symbol = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))
                    indexOfVar = tmpRawStockDailyData.IndexOf("open")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 7)
                    tmpDailyStockQuote.openPrice = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))
                    indexOfVar = tmpRawStockDailyData.IndexOf("high")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 7)
                    tmpDailyStockQuote.highPrice = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))
                    indexOfVar = tmpRawStockDailyData.IndexOf("low")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 6)
                    tmpDailyStockQuote.lowPrice = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))

                    indexOfVar = tmpRawStockDailyData.IndexOf("ltp")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 6)
                    tmpDailyStockQuote.lastTradedPrice = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))
                    indexOfVar = tmpRawStockDailyData.IndexOf("ptsC")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 7)
                    tmpDailyStockQuote.changeinPrice = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))
                    indexOfVar = tmpRawStockDailyData.IndexOf("per")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 6)
                    tmpDailyStockQuote.changeInPercentage = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))
                    indexOfVar = tmpRawStockDailyData.IndexOf("trdVol")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 9)
                    tmpDailyStockQuote.volume = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))

                    indexOfVar = tmpRawStockDailyData.IndexOf("ntp")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 6)
                    tmpDailyStockQuote.turnover = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))
                    indexOfVar = tmpRawStockDailyData.IndexOf("wkhi")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 8)
                    tmpDailyStockQuote.yearlyHighPrice = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))
                    indexOfVar = tmpRawStockDailyData.IndexOf("wklo")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 7)
                    tmpDailyStockQuote.yearlyLowPrice = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf("""}"))
                    indexOfVar = tmpRawStockDailyData.IndexOf("yPC")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 6)
                    tmpDailyStockQuote.yearlyPercentageChange = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))
                    indexOfVar = tmpRawStockDailyData.IndexOf("mPC")
                    tmpRawStockDailyData = tmpRawStockDailyData.Substring(indexOfVar + 6)
                    tmpDailyStockQuote.monthlyPercentageChange = tmpRawStockDailyData.Substring(0, tmpRawStockDailyData.IndexOf(""","))

                    stockList.Add(tmpDailyStockQuote.symbol)
                    dailyStockDetailsList.Add(tmpDailyStockQuote)
                End If

            End While
        Catch ex As Exception
            StockAppLogger.LogError("getDailyStockDetails Error in creating daily record ", ex, "DailyStockQuote")
        End Try
        StockAppLogger.Log("getDailyStockDetails End", "DailyStockQuote")
    End Sub

    Private Sub StoreDailyStockDetail()
        Dim insertStatement As String
        Dim insertValues As String

        StockAppLogger.Log("StoreDailyStockDetail Start", "DailyStockQuote")
        Try {
            For Each tmpDailyStockDetails In dailyStockDetailsList
                insertStatement = "INSERT INTO DAILYSTOCKDATA (STOCKNAME, OPENPRICE, HIGHPRICE, LOWPRICE, LAST_TRADED_PRICE, CHANGE, CHANGE_PERCENTAGE, VOLUME, TURNOVER_IN_CRS, YEARLY_HIGH, YEARLY_LOW, YERLY_PERCENTAGE_CHANGE, MONTHLY_PERCENTAGE_CHANGE, UPDATEDATE, UPDATETIME)"
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
                insertValues = insertValues & "'" & tmpDailyStockDetails.updateDate & "',"
                insertValues = insertValues & "'" & tmpDailyStockDetails.updateTime & "');"
                DBFunctions.ExecuteSQLStmt(insertStatement & insertValues)
            Next

            DBFunctions.CloseSQLConnection()
        Catch ex As Exception
            StockAppLogger.LogError("StoreDailyStockDetail Error in storing daily record", ex, "DailyStockQuote")
        End Try
        StockAppLogger.Log("StoreDailyStockDetail End", "DailyStockQuote")
    End Sub

End Class
