Imports System.Collections.Generic

Public Class NSEindices
    Dim indexName As String
    Dim lastPrice As Double
    Dim priceChange As Double
    Dim percentageChange As Double
    Dim priceDate As Date
    'Dim StockAppLogger As StockAppLogger = StockAppLogger.InitializeLogger("NSEindices")

    Public Function getIndicesListAndStore() As Boolean
        Dim rawIndicesData As String
        Dim executionresult As Boolean

        StockAppLogger.Log("getIndicesListAndStore Start")
        rawIndicesData = Helper.GetDataFromUrl(My.Settings.NSEIndicesURL)
        Dim NSEindicesList As List(Of NSEindices) = parseAndPopulateObjects(rawIndicesData)
        executionresult = storeIndicesDatainDB(NSEindicesList)
        StockAppLogger.Log("getIndicesListAndStore End")
        Return executionresult
    End Function

    Private Function parseAndPopulateObjects(ByVal rawIndicesData As String) As List(Of NSEindices)
        Dim NSEindicesList As List(Of NSEindices) = New List(Of NSEindices)
        Dim tmpRawIndicesData As String
        Dim indexOfVar As Integer
        Dim NSEIndicesData As NSEindices
        Dim countOfSymbols As Integer

        StockAppLogger.Log("parseAndPopulateObjects Start")
        tmpRawIndicesData = rawIndicesData
        countOfSymbols = rawIndicesData.Split("{""name").Length - 2
        For count = 1 To countOfSymbols
            NSEIndicesData = New NSEindices()
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
        StockAppLogger.Log("parseAndPopulateObjects End")
        Return NSEindicesList
    End Function

    Private Function storeIndicesDatainDB(NSEindicesList As List(Of NSEindices)) As Boolean
        Dim insertStatement As String
        Dim insertValues As String
        Dim myDataLayer As DataLayer = New DataLayer

        StockAppLogger.Log("storeIndicesDatainDB Start")
        insertStatement = "INSERT INTO NSEINDICES (INDEX_NAME, LAST_PRICE, PRICE_CHANGE, PERCENTAGE_CHANGE, PRICE_DATE)"

        For Each tmpNSEIndices In NSEindicesList
            Try
                insertValues = "VALUES ("
                insertValues = insertValues + "'" + tmpNSEIndices.indexName + "',"
                insertValues = insertValues + tmpNSEIndices.lastPrice.ToString("R") + ","
                insertValues = insertValues + tmpNSEIndices.priceChange.ToString("R") + ","
                insertValues = insertValues + tmpNSEIndices.percentageChange.ToString("R") + ","
                insertValues = insertValues + "'" + tmpNSEIndices.priceDate + "');"
                DBFunctions.ExecuteSQLStmtExt(insertStatement + insertValues, "DC")
                'myDataLayer.ExecuteSQLStmt(insertStatement + insertValues)
            Catch exc As Exception
                StockAppLogger.LogError("storeIndicesDatainDB Error Occurred in inserting IndicesList for index = " & tmpNSEIndices.indexName, exc)
            End Try
        Next
        DBFunctions.CloseSQLConnectionExt("DC")
        StockAppLogger.Log("storeIndicesDatainDB End")
        Return True
    End Function
End Class
