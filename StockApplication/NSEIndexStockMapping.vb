Imports System.Collections.Generic

Public Class NSEIndexStockMapping

    Dim indexName As String
    Dim stockName As String
    Dim myLogger As StockAppLogger = StockAppLogger.InitializeLogger("NSEIndexStockMapping")

    Public Function createIndicesToStockMapping(ByVal rawIndicestoStockMapping As String, ByVal indexName As String) As Boolean
        Dim objNSEIndexStockMapping As List(Of NSEIndexStockMapping)
        myLogger.Log("createIndicesToStockMapping Strt")
        objNSEIndexStockMapping = populateObjectFromRawData(rawIndicestoStockMapping, indexName)
        storeIndicesToStockMapping(objNSEIndexStockMapping)
        myLogger.Log("createIndicesToStockMapping End")
        Return True
    End Function

    Private Function populateObjectFromRawData(ByVal rawIndicestoStockMapping As String, ByVal indexName As String) As List(Of NSEIndexStockMapping)
        Dim objNSEIndexStockMapping As List(Of NSEIndexStockMapping) = New List(Of NSEIndexStockMapping)
        Dim tmpRawIndicesData As String
        Dim tmpNSEIndexStockMapping As NSEIndexStockMapping
        Dim countOfSymbols As Integer
        Dim indexOfVar As Integer

        myLogger.Log("populateObjectFromRawData Start")
        tmpRawIndicesData = rawIndicestoStockMapping
        countOfSymbols = rawIndicestoStockMapping.Split("{""symbol").Length - 1
        For count = 1 To countOfSymbols
            tmpNSEIndexStockMapping = New NSEIndexStockMapping()
            tmpNSEIndexStockMapping.indexName = indexName
            indexOfVar = tmpRawIndicesData.IndexOf("symbol")
            tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 9)
            tmpNSEIndexStockMapping.stockName = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
            objNSEIndexStockMapping.Add(tmpNSEIndexStockMapping)
        Next
        myLogger.Log("populateObjectFromRawData End")
        Return objNSEIndexStockMapping
    End Function

    Private Function storeIndicesToStockMapping(objNSEIndexStockMapping As List(Of NSEIndexStockMapping)) As Boolean

        Dim updateOrInsert As String
        Dim updateOrInsertValues As String

        myLogger.Log("storeIndicesToStockMapping Start")
        updateOrInsert = "update Or insert into NSE_INDICES_TO_STOCK_MAPPING (INDEX_NAME, STOCK_NAME) values("
        For Each tmpNSEIndexStockMapping In objNSEIndexStockMapping
            updateOrInsertValues = "'" & tmpNSEIndexStockMapping.indexName & "', '" & tmpNSEIndexStockMapping.stockName & "') matching(INDEX_NAME, STOCK_NAME);"
            DBFunctions.ExecuteSQLStmt(updateOrInsert & updateOrInsertValues)
        Next
        DBFunctions.CloseSQLConnection()
        myLogger.Log("storeIndicesToStockMapping End")
        Return True
    End Function
End Class
