Imports System.Collections.Generic
Public Class NSEindices
    Dim indexName As String
    Dim lastPrice As Double
    Dim priceChange As Double
    Dim percentageChange As Double
    Dim priceDate As Date
    Dim myLogger As StockAppLogger = StockAppLogger.InitializeLogger()

    Public Function getIndicesListAndStore() As Boolean
        Dim rawIndicesData As String
        Dim executionresult As Boolean

        myLogger.Log("getIndicesListAndStore Start")
        rawIndicesData = Helper.GetDataFromUrl("https://www.nseindia.com/homepage/Indices1.json")
        Dim NSEindicesList As List(Of NSEindices) = parseAndPopulateObjects(rawIndicesData)
        executionresult = storeIndicesDatainDB(NSEindicesList)
        myLogger.Log("getIndicesListAndStore End")
        Return executionresult
    End Function

    Private Function parseAndPopulateObjects(ByVal rawIndicesData As String) As List(Of NSEindices)
        Dim NSEindicesList As List(Of NSEindices) = New List(Of NSEindices)
        Dim tmpRawIndicesData As String
        Dim indexOfVar As Integer
        Dim NSEIndicesData As NSEindices
        Dim countOfSymbols As Integer

        myLogger.Log("parseAndPopulateObjects Start")
        tmpRawIndicesData = rawIndicesData
        countOfSymbols = rawIndicesData.Split("{""name").Length - 2
        For count = 1 To countOfSymbols
            NSEIndicesData = New NSEindices
            indexOfVar = tmpRawIndicesData.IndexOf("name")
            tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 7)
            NSEIndicesData.indexName = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
            indexOfVar = tmpRawIndicesData.IndexOf("lastPrice")
            tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 12)
            NSEIndicesData.lastPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))

            indexOfVar = tmpRawIndicesData.IndexOf("change")
            tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 9)
            NSEIndicesData.priceChange = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))

            indexOfVar = tmpRawIndicesData.IndexOf("pChange")
            tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 10)
            NSEIndicesData.percentageChange = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))

            NSEIndicesData.priceDate = Today
            NSEindicesList.Add(NSEIndicesData)
        Next count
        myLogger.Log("parseAndPopulateObjects End")
        Return NSEindicesList
    End Function

    Private Function storeIndicesDatainDB(NSEindicesList As List(Of NSEindices)) As Boolean
        Dim insertStatement As String
        Dim insertValues As String
        Dim myDataLayer As DataLayer = New DataLayer

        myLogger.Log("storeIndicesDatainDB Start")
        insertStatement = "INSERT INTO NSEINDICES (INDEX_NAME, LAST_PRICE, PRICE_CHANGE, PERCENTAGE_CHANGE, PRICE_DATE)"

        For Each tmpNSEIndices In NSEindicesList
            Try
                insertValues = "VALUES ("
                insertValues = insertValues + "'" + tmpNSEIndices.indexName + "',"
                insertValues = insertValues + tmpNSEIndices.lastPrice.ToString("R") + ","
                insertValues = insertValues + tmpNSEIndices.priceChange.ToString("R") + ","
                insertValues = insertValues + tmpNSEIndices.percentageChange.ToString("R") + ","
                insertValues = insertValues + "'" + tmpNSEIndices.priceDate + "');"
                myDataLayer.ExecuteSQLStmt(insertStatement + insertValues)
            Catch exc As Exception
                myLogger.LogError("Error Occurred in inserting IndicesList = ", exc)
            End Try
        Next
        myLogger.Log("storeIndicesDatainDB End")
        Return True
    End Function
End Class
