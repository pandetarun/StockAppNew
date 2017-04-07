Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class BollingerBands

    Dim lastTradedPrice, BBUper, BBLower, simpleMA As Double
    Dim PeriodBandwidth As Double
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

    Private Function IntraDayBBCalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim whereClause As String
        Dim orderClause As String
        Dim closingPrice As Double
        Dim tradedPrice, totalTradedPrice As Double
        Dim perioddeviation As Double
        Dim PeriodData As List(Of Double) = New List(Of Double)

        Dim tmpPeriodData As List(Of Double)

        StockAppLogger.Log("IntraDayBBCalculation Start", "BollingerBands")

        lastTradedPrice = 0
        counter = 0
        totalTradedPrice = 0
        simpleMA = 0
        BBUper = 0
        BBLower = 0
        MAStock = tmpStockCode
        Try
            whereClause = "TRADEDDATE='" & Today & "' and companycode = '" & tmpStockCode & "'"
            orderClause = "lastupdatetime desc"
            MADate = Today
            ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " lastClosingPrice, lastupdatetime", whereClause, orderClause)
            While ds.Read()
                counter = counter + 1
                If counter = 1 Then
                    MATime = ds.GetValue(ds.GetOrdinal("lastupdatetime"))
                    closingPrice = ds.GetValue(ds.GetOrdinal("lastClosingPrice"))
                    lastTradedPrice = closingPrice
                End If
                tradedPrice = Double.Parse(ds.GetValue(ds.GetOrdinal("lastClosingPrice")))
                totalTradedPrice = totalTradedPrice + tradedPrice
                PeriodData.Add(tradedPrice)
                If counter = 14 Then
                    perioddeviation = 0
                    BBLower = 0
                    BBUper = 0
                    tmpPeriodData = New List(Of Double)
                    simpleMA = totalTradedPrice / 14
                    For counter1 As Integer = 1 To 14
                        tmpPeriodData.Item(counter1) = PeriodData.Item(counter1) - simpleMA
                        tmpPeriodData.Item(counter1) = tmpPeriodData.Item(counter1) * tmpPeriodData.Item(counter1)
                        perioddeviation = perioddeviation + tmpPeriodData.Item(counter1)
                    Next counter1
                    perioddeviation = perioddeviation / 14
                    perioddeviation = Math.Sqrt(perioddeviation)
                    BBLower = simpleMA - 2 * perioddeviation
                    BBUper = simpleMA + 2 * perioddeviation
                    PeriodBandwidth = BBUper - BBLower
                    InsertIntraDayBBtoDB(14)
                End If

                If counter = 20 Then
                    perioddeviation = 0
                    BBLower = 0
                    BBUper = 0
                    tmpPeriodData = New List(Of Double)
                    simpleMA = totalTradedPrice / 20
                    For counter1 As Integer = 1 To 20
                        tmpPeriodData.Item(counter1) = PeriodData.Item(counter1) - simpleMA
                        tmpPeriodData.Item(counter1) = tmpPeriodData.Item(counter1) * tmpPeriodData.Item(counter1)
                        perioddeviation = perioddeviation + tmpPeriodData.Item(counter1)
                    Next counter1
                    perioddeviation = perioddeviation / 20
                    perioddeviation = Math.Sqrt(perioddeviation)
                    BBLower = simpleMA - 2 * perioddeviation
                    BBUper = simpleMA + 2 * perioddeviation
                    PeriodBandwidth = BBUper - BBLower
                    InsertIntraDayBBtoDB(20)
                End If
                If counter = 26 Then
                    perioddeviation = 0
                    BBLower = 0
                    BBUper = 0
                    tmpPeriodData = New List(Of Double)
                    simpleMA = totalTradedPrice / 26
                    For counter1 As Integer = 1 To 26
                        tmpPeriodData.Item(counter1) = PeriodData.Item(counter1) - simpleMA
                        tmpPeriodData.Item(counter1) = tmpPeriodData.Item(counter1) * tmpPeriodData.Item(counter1)
                        perioddeviation = perioddeviation + tmpPeriodData.Item(counter1)
                    Next counter1
                    perioddeviation = perioddeviation / 26
                    perioddeviation = Math.Sqrt(perioddeviation)
                    BBLower = simpleMA - 2 * perioddeviation
                    BBUper = simpleMA + 2 * perioddeviation
                    PeriodBandwidth = BBUper - BBLower
                    InsertIntraDayBBtoDB(26)
                    Exit While
                End If
            End While
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in calculating intraday Bollinger Band = ", exc, "BollingerBands")
            Return False
        End Try
        StockAppLogger.Log("IntraDayBBCalculation End", "BollingerBands")
        Return True
    End Function

    Private Function InsertIntraDayBBtoDB(ByVal period As Integer) As Boolean

        Dim insertStatement As String
        Dim insertValues As String
        Dim sqlStatement As String
        Dim fireQuery As Boolean = False

        StockAppLogger.Log("InsertIntraDayBBtoDB Start", "BollingerBands")
        Try
            insertStatement = "INSERT INTO INTRADAYBOLLINGERBANDS (TRADEDDATE, STOCK_NAME, LASTUPDATETIME, PERIOD, CLOSINGPRICE, SMA, UPPERBAND, LOWERBAND, BANDWIDTH"
            insertValues = "VALUES ('" & MADate & "', '" & MAStock & "', '" & MATime & "', " & period & ", " & lastTradedPrice & ", " & simpleMA & ", " & BBUper & ", " & BBLower & ", " & PeriodBandwidth & ")"

            sqlStatement = insertStatement & insertValues
            DBFunctions.ExecuteSQLStmt(sqlStatement)
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in storing intraday bollinger band = ", exc, "BollingerBands")
            Return False
        End Try
        StockAppLogger.Log("InsertIntraDayBBtoDB End", "BollingerBands")
        Return True
    End Function
End Class

