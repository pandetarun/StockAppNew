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
        Dim tmpRawIndicesData As String
        Dim indexOfVar As Integer
        'Dim NSEIndicesData As NSEindices
        ' Dim countOfSymbols As Integer

        myLogger.Log("getNSEIndicesDetails Start")
        rawIndicesData = Helper.GetDataFromUrl(NSEIndicesURL)
        tmpRawIndicesData = rawIndicesData
        'countOfSymbols = rawIndicesData.Split("{""name").Length - 2
        'For count = 1 To countOfSymbols
        'NSEIndicesData = New NSEindices
        indexOfVar = tmpRawIndicesData.IndexOf("indexName")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 11)
        tmpNSEIndicesDetails.symbol = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("open")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 6)
        tmpNSEIndicesDetails.openPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("high")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 6)
        tmpNSEIndicesDetails.highPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("low")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 5)
        tmpNSEIndicesDetails.lowPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))

        indexOfVar = tmpRawIndicesData.IndexOf("ltp")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 5)
        tmpNSEIndicesDetails.lastTradedPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("ch")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 4)
        tmpNSEIndicesDetails.changeinPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("per")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 5)
        tmpNSEIndicesDetails.changeInPercentage = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("yCls")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 6)
        tmpNSEIndicesDetails.yearlyPercentageChange = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("mCls")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 6)
        tmpNSEIndicesDetails.monthlyPercentageChange = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("yHigh")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 7)
        tmpNSEIndicesDetails.yearlyHighPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("yLow")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 6)
        tmpNSEIndicesDetails.yearlyLowPrice = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))
        indexOfVar = tmpRawIndicesData.IndexOf("trdValueSum")
        tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 13)
        tmpNSEIndicesDetails.turnover = tmpRawIndicesData.Substring(0, tmpRawIndicesData.IndexOf(""","))

        indexOfVar = tmpRawIndicesData.IndexOf("trdVolumesum")
        'tmpRawIndicesData = tmpRawIndicesData.Substring(indexOfVar + 12)
        tmpNSEIndicesDetails.monthlyPercentageChange = tmpRawIndicesData.Substring(indexOfVar + 12).Substring(0, tmpRawIndicesData.IndexOf(""","))
        NSEIndicesData.priceDate = Today
        NSEindicesList.Add(NSEIndicesData)
        'Next count
        myLogger.Log("getNSEIndicesDetails End")
    End Function

End Class
