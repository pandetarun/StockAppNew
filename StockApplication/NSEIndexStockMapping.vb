Imports System.Collections.Generic

Public Class NSEIndexStockMapping

    Dim indexName As String
    Dim stockName As String
    Dim myLogger As StockAppLogger = StockAppLogger.InitializeLogger("NSEIndexStockMapping")

    Public Function createIndicesToStockMapping(ByVal rawIndicestoStockMapping As String, ByVal indexName As String) As Boolean
        Dim objNSEIndexStockMapping As List(Of NSEIndexStockMapping)

        objNSEIndexStockMapping = populateObjectFromRawData(rawIndicestoStockMapping, indexName)

        Return True
    End Function

    Public Function populateObjectFromRawData(ByVal rawIndicestoStockMapping As String, ByVal indexName As String) As List(Of NSEIndexStockMapping)
        Dim objNSEIndexStockMapping As List(Of NSEIndexStockMapping) = New List(Of NSEIndexStockMapping)
        Dim tmpRawIndicesData As String
        Dim tmpNSEIndexStockMapping As NSEIndexStockMapping
        Dim countOfSymbols As Integer
        Dim indexOfVar As Integer

        myLogger.Log("parseAndPopulateObjects Start")
        tmpRawIndicesData = rawIndicestoStockMapping
        countOfSymbols = rawIndicestoStockMapping.Split("{""symbol").Length - 2
        For count = 1 To countOfSymbols
            tmpNSEIndexStockMapping = New NSEIndexStockMapping()
            tmpNSEIndexStockMapping.indexName = indexName
            indexOfVar = tmpRawIndicesData.IndexOf("symbol")
            tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 8)
            tmpNSEIndexStockMapping.stockName = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
            objNSEIndexStockMapping.Add(tmpNSEIndexStockMapping)
        Next
        Return objNSEIndexStockMapping
    End Function

    Public Function storeIndicesToStockMapping(objNSEIndexStockMapping As List(Of NSEIndexStockMapping)) As Boolean

        Dim updateOrInsert As String
        Dim updateOrInsertValues As String

        updateOrInsert = "update Or insert into NSE_INDICES_TO_STOCK_MAPPING (INDEX_NAME, STOCK_NAME) values("

        'Suzy Creamcheese', 3278823, 'Green Pastures') matching(Number) returning rec_id into : id;

        For Each tmpNSEIndexStockMapping In objNSEIndexStockMapping
            updateOrInsertValues = "'" & tmpNSEIndexStockMapping.indexName & "', '" & tmpNSEIndexStockMapping.stockName & "') matching(INDEX_NAME, STOCK_NAME);"
            DBFunctions.ExecuteSQLStmt(updateOrInsert & updateOrInsertValues)
        Next
        DBFunctions.CloseSQLConnection()
        Return True
    End Function

End Class
