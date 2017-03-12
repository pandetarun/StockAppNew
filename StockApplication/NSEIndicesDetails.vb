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
    Dim myLogger As StockAppLogger = StockAppLogger.InitializeLogger("NSEIndicesDetails")

    Dim pathOfFileofURLs As String = My.Settings.ApplicationFileLocation & "\URLs.txt"
    Dim RawNSEindicesDetailsList As List(Of String) = New List(Of String)
    Dim NSEindicesDetailsList As List(Of NSEIndicesDetails) = New List(Of NSEIndicesDetails)

    Public Function getIndicesDetailsAndStore() As Boolean
        Dim NSEindicesUrlsList As List(Of String) = New List(Of String)
        myLogger.Log("getIndicesDetailsAndStore Start")
        NSEindicesUrlsList = GetNSEUrlForAllIndices()
        getAllIndicesDetails(NSEindicesUrlsList)

        myLogger.Log("getIndicesDetailsAndStore End")
        Return True
    End Function

    Private Function GetNSEUrlForAllIndices() As List(Of String)
        Dim NSEindicesUrlsList As List(Of String) = New List(Of String)
        Dim reader = File.OpenText(pathOfFileofURLs)
        Dim line As String = Nothing
        myLogger.Log("GetNSEUrlForAllIndices Start")
        While (reader.Peek() <> -1)
            line = reader.ReadLine()
            NSEindicesUrlsList.Add(line)
        End While
        myLogger.Log("GetNSEUrlForAllIndices End")
        Return NSEindicesUrlsList
    End Function

    Private Function getAllIndicesDetails(ByVal NSEindicesUrlsList As List(Of String)) As List(Of String)
        Dim NSEindicesDetailsList As List(Of String) = New List(Of String)
        myLogger.Log("getAllIndicesDetails Start")
        For Each NSEIndicesURL In NSEindicesUrlsList
            getNSEIndicesDetails(NSEIndicesURL)
        Next
        myLogger.Log("getAllIndicesDetails End")
        Return NSEindicesDetailsList
    End Function

    Private Function getNSEIndicesDetails(ByVal NSEIndicesURL As String)
        Dim rawIndicesData As String
        Dim tmpNSEIndicesDetails As NSEIndicesDetails = New NSEIndicesDetails()


        Dim indexOfVar As Integer
        Dim NSEIndicesData As NSEindices
        Dim countOfSymbols As Integer

        myLogger.Log("getNSEIndicesDetails Start")
        rawIndicesData = Helper.GetDataFromUrl(NSEIndicesURL)
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
        myLogger.Log("getNSEIndicesDetails End")
    End Function

End Class
