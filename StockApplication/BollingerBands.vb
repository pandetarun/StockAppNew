Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class BollingerBands

    Dim lastTradedPrice, fourteenSampleBBUper, fourteenSampleBBLower, twentySampleBBUper, twentySampleBBLower, twentySixSampleBBUper, twentySixSampleBBLower, simpleMA As Double
    Dim MATime, MAStock As String
    Dim MADate As Date
    Dim counter As Integer

    Public Function CalculateAndStoreBollingerBands() As Boolean
        'Dim tmpStockList As List(Of String) = New List(Of String)
        GetStockListAndCalculateBollingerBands()
        Return False
    End Function

    Private Function GetStockListAndCalculateBollingerBands() As Boolean
        Dim tmpStockList As List(Of String) = New List(Of String)
        Dim tmpStockCode As String
        Dim ds As FbDataReader = Nothing

        StockAppLogger.Log("GetStockListAndCalculateBollingerBands Start", "BollingerBands")
        Try
            ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING")
            While ds.Read()
                tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
                If Not tmpStockList.Contains(tmpStockCode) Then
                    tmpStockList.Add(tmpStockCode)
                    'If IntraDayMACalculation(tmpStockCode) Then
                    'InsertIntraDayMAtoDB()
                    'Else
                    'StockAppLogger.Log("GetStockListAndCalculateBollingerBands BollingerBands failed for Stock = " & tmpStockCode & " at time = " & MATime, "MovingAverage")
                    'End If
                End If
            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting stocklist from DB = ", exc, "BollingerBands")
            Return False
        End Try
        StockAppLogger.Log("GetStockListAndCalculateBollingerBands End", "BollingerBands")
        Return True
    End Function

    Private Function IntraDayMACalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim whereClause, whereClause1 As String
        Dim orderClause, orderClause1 As String
        Dim closingPrice As Double
        Dim laststoredEMA As Double
        Dim recordPresentInTable As Boolean

        StockAppLogger.Log("IntraDayMACalculation Start", "MovingAverage")

        lastTradedPrice = 0
        counter = 0
        simpleMA = 0
        fourteenSampleBBUper = 0
        fourteenSampleBBLower = 0
        twentySampleBBUper = 0
        twentySampleBBLower = 0
        twentySixSampleBBUper = 0
        twentySixSampleBBLower = 0

        MAStock = tmpStockCode

        Try
            whereClause1 = "TRADEDDATE='" & Today & "' and STOCK_NAME = '" & tmpStockCode & "'"
            orderClause1 = "LASTUPDATETIME desc"
            whereClause = "TRADEDDATE='" & Today & "' and companycode = '" & tmpStockCode & "'"
            orderClause = "lastupdatetime desc"

            MADate = Today
            'orderClause1 = "lastupdatetime"
            ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " lastClosingPrice, lastupdatetime", whereClause, orderClause)
            ds1 = DBFunctions.getDataFromTable("INTRADAYMOVINGAVERAGES", " * ", whereClause1, orderClause1)
            recordPresentInTable = ds1.Read()
            While ds.Read()
                counter = counter + 1
                If counter = 1 Then
                    MATime = ds.GetValue(ds.GetOrdinal("lastupdatetime"))
                    closingPrice = ds.GetValue(ds.GetOrdinal("lastClosingPrice"))
                    lastTradedPrice = closingPrice
                End If
                simpleMA = simpleMA + Double.Parse(ds.GetValue(ds.GetOrdinal("lastClosingPrice")))
                If counter = 3 Then
                    'threeSampleSMA = simpleMA / 3
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("THREEEMA")))
                        'threeSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        'threeSampleEMA = threeSampleSMA
                    End If

                ElseIf counter = 20 Then
                    'twentySampleSMA = simpleMA / 20
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("TWENTYEMA")))
                        ' twentySampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        'twentySampleEMA = twentySampleSMA
                    End If
                ElseIf counter = 50 Then
                    ' FiftySampleSMA = simpleMA / 50
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIFTYEMA")))
                        'FiftySampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        'FiftySampleEMA = FiftySampleSMA
                    End If
                    Exit While
                End If
            End While
            If counter = 1 Then
                'threeSampleSMA = simpleMA

            End If
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in calculating intraday moving average = ", exc, "MovingAverage")
            Return False
        End Try
        StockAppLogger.Log("IntraDayMACalculation End", "MovingAverage")
        Return True
    End Function
End Class
