Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic


Public Class CalculateTechnicalIndicators

    Public Sub CalculateIntradayIndicators()
        Dim tmpStockList As List(Of String) = New List(Of String)
        Dim tmpStockCode As String
        Dim ds As FbDataReader = Nothing
        Dim tmpMovingAverageCalculation As MovingAverage = New MovingAverage()
        Dim tmpBollingerBandCalculation As BollingerBands = New BollingerBands()
        Dim tmpMACDCalculation As MACD = New MACD()

        StockAppLogger.Log("CalculateIntradayIndicators Start", "CalculateTechnicalIndicators")
        Try
            ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING")
            While ds.Read()
                tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
                If Not tmpStockList.Contains(tmpStockCode) Then
                    tmpStockList.Add(tmpStockCode)
                    Try
                        'Moving Average calculation
                        If tmpMovingAverageCalculation.IntraDaySNEMACalculation(tmpStockCode) Then
                            StockAppLogger.Log("CalculateIntradayIndicators intraday moving average successfull for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                        Else
                            StockAppLogger.LogInfo("CalculateIntradayIndicators intraday moving average failed for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                        End If
                    Catch exc As Exception
                        StockAppLogger.LogError("CalculateIntradayIndicators Error Occurred in calculating intraday moving average = ", exc, "CalculateTechnicalIndicators")
                    End Try

                    Try
                        'Intra day Bollinger Band calculation
                        If tmpBollingerBandCalculation.IntraDayBBCalculation(tmpStockCode) Then
                            StockAppLogger.LogInfo("CalculateIntradayIndicators BollingerBands stored for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                        Else
                            StockAppLogger.LogInfo("CalculateIntradayIndicators BollingerBands failed for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                        End If
                    Catch exc As Exception
                        StockAppLogger.LogError("CalculateIntradayIndicators Error Occurred in calculating bollinger band = ", exc, "CalculateTechnicalIndicators")
                    End Try

                    Try
                        'Intra day MACD calculation
                        If tmpMACDCalculation.IntraDayMACDCalculation(tmpStockCode) Then
                            StockAppLogger.Log("CalculateIntradayIndicators MACD successfull for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                        Else
                            StockAppLogger.LogInfo("CalculateIntradayIndicators MACD failed for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                        End If
                    Catch ex As Exception
                        StockAppLogger.LogError("CalculateIntradayIndicators Error Occurred in calculating MACD = ", ex, "CalculateTechnicalIndicators")
                    End Try
                End If
            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("CalculateIntradayIndicators Error Occurred in calculating intraday indicator = ", exc, "CalculateTechnicalIndicators")
            'Return False
        End Try
        StockAppLogger.Log("CalculateIntradayIndicators End", "CalculateTechnicalIndicators")
    End Sub

    Public Sub CalculateDailyIndicators()
        Dim tmpStockList As List(Of String) = New List(Of String)
        Dim tmpStockCode As String
        Dim ds As FbDataReader = Nothing
        Dim tmpMovingAverageCalculation As MovingAverage = New MovingAverage()
        Dim tmpBollingerBandCalculation As BollingerBands = New BollingerBands()
        Dim tmpMACDCalculation As MACD = New MACD()

        StockAppLogger.Log("CalculateDailyIndicators Start", "CalculateTechnicalIndicators")
        Try
            ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING")
            While ds.Read()
                tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
                If Not tmpStockList.Contains(tmpStockCode) Then
                    tmpStockList.Add(tmpStockCode)
                    Try
                        'Moving Average calculation
                        If tmpMovingAverageCalculation.DailySNEMACalculation(tmpStockCode) Then
                            StockAppLogger.Log("CalculateDailyIndicators daily moving average successfull for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                        Else
                            StockAppLogger.LogInfo("CalculateDailyIndicators daily moving average failed for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                        End If
                    Catch exc As Exception
                        StockAppLogger.LogError("CalculateDailyIndicators Error Occurred in calculating daily moving average = ", exc, "CalculateTechnicalIndicators")
                    End Try

                    Try
                        'daily Bollinger Band calculation
                        If tmpBollingerBandCalculation.DailyBBCalculation(tmpStockCode) Then
                            StockAppLogger.LogInfo("CalculateDailyIndicators BollingerBands stored for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                        Else
                            StockAppLogger.LogInfo("CalculateDailyIndicators BollingerBands failed for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                        End If
                    Catch exc As Exception
                        StockAppLogger.LogError("CalculateDailyIndicators Error Occurred in calculating daily bollinger band = ", exc, "CalculateTechnicalIndicators")
                    End Try
                End If
            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("CalculateDailyIndicators Error Occurred in calculating daily indicator = ", exc, "CalculateTechnicalIndicators")
            'Return False
        End Try
        StockAppLogger.Log("CalculateDailyIndicators End", "CalculateTechnicalIndicators")
    End Sub

End Class
