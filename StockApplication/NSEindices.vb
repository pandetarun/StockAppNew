Imports System.Collections.Generic
Public Class NSEindices
    Dim indexName As String
    Dim lastPrice As Double
    Dim priceChange As Double
    Dim percentageChange As Double

    Public Function getIndicesListAndStore() As Boolean
        Dim rawIndicesData As String

        rawIndicesData = Helper.GetDataFromUrl("https://www.nseindia.com/homepage/Indices1.json")
        Dim NSEindicesList As List(Of NSEindices) = parseAndPopulateObjects(rawIndicesData)
        Return storeIndicesDatainDB(NSEindicesList)
    End Function

    Private Function parseAndPopulateObjects(ByVal rawIndicesData As String) As List(Of NSEindices)
        Dim NSEindicesList As List(Of NSEindices) = New List(Of NSEindices)
        Dim tmpRawIndicesData As String
        Dim indexOfVar As Integer
        Dim NSEIndicesData As NSEindices
        Dim countOfSymbols As Integer

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
            NSEindicesList.Add(NSEIndicesData)
        Next count
        Return NSEindicesList
    End Function

    Private Function storeIndicesDatainDB(NSEindicesList As List(Of NSEindices)) As Boolean

        Return True
    End Function

End Class
