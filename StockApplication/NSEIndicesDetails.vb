Imports System.Collections.Generic
Imports System
Imports System.IO

Public Class NSEIndicesDetails

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
    'Dim StockAppLogger As StockAppLogger = StockAppLogger.InitializeLogger("NSEIndicesDetails")

    Dim pathOfFileofURLs As String = My.Settings.ApplicationFileLocation & "\URLs.txt"
    Dim RawNSEindicesDetailsList As List(Of String) = New List(Of String)
    Dim NSEindicesDetailsList As List(Of NSEIndicesDetails) = New List(Of NSEIndicesDetails)

    Public Function getIndicesDetailsAndStore() As Boolean
        Dim NSEindicesUrlsList As List(Of String)
        StockAppLogger.Log("getIndicesDetailsAndStore Start")
        NSEindicesUrlsList = GetNSEUrlForAllIndices()
        getAllIndicesDetails(NSEindicesUrlsList)

        StockAppLogger.Log("getIndicesDetailsAndStore End")
        Return True
    End Function

    Private Function GetNSEUrlForAllIndices() As List(Of String)
        Dim NSEindicesUrlsList As List(Of String) = New List(Of String)
        Dim reader = File.OpenText(pathOfFileofURLs)
        Dim line As String = Nothing
        StockAppLogger.Log("GetNSEUrlForAllIndices Start")
        While (reader.Peek() <> -1)
            line = reader.ReadLine()
            NSEindicesUrlsList.Add(line)
        End While
        StockAppLogger.Log("GetNSEUrlForAllIndices End")
        Return NSEindicesUrlsList
    End Function

    Private Function getAllIndicesDetails(ByVal NSEindicesUrlsList As List(Of String)) As List(Of String)
        Dim NSEindicesDetailsList As List(Of String) = New List(Of String)

        StockAppLogger.Log("getAllIndicesDetails Start")
        For Each NSEIndicesURL In NSEindicesUrlsList
            getNSEIndicesDetails(NSEIndicesURL)
            StoreOrUpdateIndicesDetail()

        Next
        StockAppLogger.Log("getAllIndicesDetails End")
        Return NSEindicesDetailsList
    End Function

    Private Sub getNSEIndicesDetails(ByVal NSEIndicesURL As String)
        Dim rawIndicesData As String
        Dim tmpNSEIndicesDetails As NSEIndicesDetails = New NSEIndicesDetails()
        Dim tmpNSEIndexStockMapping As NSEIndexStockMapping = New NSEIndexStockMapping
        Dim tmpRawIndicesData As String
        Dim indexOfVar As Integer
        Dim tmpString As String
        'Dim NSEIndicesData As NSEindices
        ' Dim countOfSymbols As Integer

        StockAppLogger.Log("getNSEIndicesDetails Start")
        rawIndicesData = Helper.GetDataFromUrl(NSEIndicesURL)
        tmpRawIndicesData = rawIndicesData

        tmpNSEIndicesDetails.updateDate = Today
        indexOfVar = tmpRawIndicesData.IndexOf("time")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 6)
        tmpString = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        tmpNSEIndicesDetails.updateTime = tmpString.Substring(tmpString.Length - 8)
        indexOfVar = tmpRawIndicesData.IndexOf("indexName")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 12)
        tmpNSEIndicesDetails.symbol = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("open")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 7)
        tmpNSEIndicesDetails.openPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("high")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 7)
        tmpNSEIndicesDetails.highPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("low")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 6)
        tmpNSEIndicesDetails.lowPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))

        indexOfVar = tmpRawIndicesData.IndexOf("ltp")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 6)
        tmpNSEIndicesDetails.lastTradedPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("ch")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 5)
        tmpNSEIndicesDetails.changeinPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("per")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 6)
        tmpNSEIndicesDetails.changeInPercentage = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("yCls")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 7)
        tmpNSEIndicesDetails.yearlyPercentageChange = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("mCls")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 7)
        tmpNSEIndicesDetails.monthlyPercentageChange = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("yHigh")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 8)
        tmpNSEIndicesDetails.yearlyHighPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("yLow")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 7)
        tmpNSEIndicesDetails.yearlyLowPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf("""}"))
        indexOfVar = tmpRawIndicesData.IndexOf("trdVolumesum")
        tmpString = tmpRawIndicesData.Substring(indexOfVar + 15)
        tmpNSEIndicesDetails.volume = tmpString.Substring(0, tmpString.IndexOf(""","))

        NSEindicesDetailsList.Add(tmpNSEIndicesDetails)
        tmpNSEIndexStockMapping.createIndicesToStockMapping(tmpRawIndicesData, tmpNSEIndicesDetails.symbol)
        StockAppLogger.Log("getNSEIndicesDetails End")
    End Sub

    Private Sub StoreOrUpdateIndicesDetail()
        Dim insertStatement As String
        Dim insertValues As String

        For Each tmpNSEIndicesDetails In NSEindicesDetailsList
            insertStatement = "INSERT INTO NSE_INDICES_DETAILS (INDEXNAME, OPENPRICE, HIGHPRICE, LOWPRICE, LAST_TRADED_PRICE, CHANGE, CHANGE_PERCENTAGE, VOLUME, TURNOVER_IN_CRS, YEARLY_HIGH, YEARLY_LOW, YERLY_PERCENTAGE_CHANGE, MONTHLY_PERCENTAGE_CHANGE, UPDATEDATE, UPDATETIME)"
            insertValues = " VALUES ('"
            insertValues = insertValues & tmpNSEIndicesDetails.symbol & "',"
            insertValues = insertValues & tmpNSEIndicesDetails.openPrice & ","
            insertValues = insertValues & tmpNSEIndicesDetails.highPrice & ","
            insertValues = insertValues & tmpNSEIndicesDetails.lowPrice & ","
            insertValues = insertValues & tmpNSEIndicesDetails.lastTradedPrice & ","
            insertValues = insertValues & tmpNSEIndicesDetails.changeinPrice & ","
            insertValues = insertValues & tmpNSEIndicesDetails.changeInPercentage & ","
            insertValues = insertValues & tmpNSEIndicesDetails.volume & ","
            insertValues = insertValues & tmpNSEIndicesDetails.turnover & ","
            insertValues = insertValues & tmpNSEIndicesDetails.yearlyHighPrice & ","
            insertValues = insertValues & tmpNSEIndicesDetails.yearlyLowPrice & ","
            insertValues = insertValues & tmpNSEIndicesDetails.yearlyPercentageChange & ","
            insertValues = insertValues & tmpNSEIndicesDetails.monthlyPercentageChange & ","
            insertValues = insertValues & "'" & tmpNSEIndicesDetails.updateDate & "',"
            insertValues = insertValues & "'" & tmpNSEIndicesDetails.updateTime & "');"
            DBFunctions.ExecuteSQLStmt(insertStatement & insertValues)
        Next
        DBFunctions.CloseSQLConnection()
    End Sub
End Class
