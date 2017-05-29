Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class GenerateIndications
    Public Sub GenerateIntraDayIndications()
        Dim tmpStockList As List(Of String) = New List(Of String)
        Dim tmpStockCode As String
        Dim ds As FbDataReader = Nothing
        Dim tmpMovingAverageCalculation As MovingAverageIndication = New MovingAverageIndication()
        'Dim tmpBollingerBandCalculation As BollingerBands = New BollingerBands()
        'Dim tmpMACDCalculation As MACD = New MACD()

        StockAppLogger.Log("GenerateIntraDayIndications Start", "GenerateIndications")
        Try
            'adasd
            ds = DBFunctions.getDataFromTableExt("NSE_INDICES_TO_STOCK_MAPPING", "DI")
            While ds.Read()
                tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
                If Not tmpStockList.Contains(tmpStockCode) Then
                    tmpStockList.Add(tmpStockCode)
                    Try
                        'Moving Average indication calculation
                        If tmpMovingAverageCalculation.GetIndicationIntraDaySMA(tmpStockCode) Then
                            StockAppLogger.LogInfo("GenerateIntraDayIndications intraday moving average indication true for Stock = " & tmpStockCode, "GenerateIndications")
                        Else
                            StockAppLogger.LogInfo("GenerateIntraDayIndications intraday moving average indication false for Stock = " & tmpStockCode, "GenerateIndications")
                        End If
                    Catch exc As Exception
                        StockAppLogger.LogError("GenerateIntraDayIndications Error Occurred in generating intraday moving average indication= ", exc, "GenerateIndications")
                    End Try

                    'Try
                    '    'Intra day Bollinger Band calculation
                    '    If tmpBollingerBandCalculation.IntraDayBBCalculation(tmpStockCode) Then
                    '        StockAppLogger.LogInfo("CalculateIntradayIndicators BollingerBands stored for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                    '    Else
                    '        StockAppLogger.LogInfo("CalculateIntradayIndicators BollingerBands failed for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                    '    End If
                    'Catch exc As Exception
                    '    StockAppLogger.LogError("CalculateIntradayIndicators Error Occurred in calculating bollinger band = ", exc, "CalculateTechnicalIndicators")
                    'End Try

                    'Try
                    '    'Intra day MACD calculation
                    '    If tmpMACDCalculation.IntraDayMACDCalculation(tmpStockCode) Then
                    '        StockAppLogger.Log("CalculateIntradayIndicators MACD successfull for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                    '    Else
                    '        StockAppLogger.LogInfo("CalculateIntradayIndicators MACD failed for Stock = " & tmpStockCode, "CalculateTechnicalIndicators")
                    '    End If
                    'Catch ex As Exception
                    '    StockAppLogger.LogError("CalculateIntradayIndicators Error Occurred in calculating MACD = ", ex, "CalculateTechnicalIndicators")
                    'End Try
                End If
            End While
            ds.Close()
            DBFunctions.CloseSQLConnectionExt("DI")
        Catch exc As Exception
            StockAppLogger.LogError("GenerateIntraDayIndications Error Occurred in generating intraday indicator = ", exc, "GenerateIndications")
            'Return False
            If ds IsNot Nothing Then
                ds.Close()
            End If
            DBFunctions.CloseSQLConnectionExt("DI")
        End Try
        StockAppLogger.Log("GenerateIntraDayIndications End", "GenerateIndications")
    End Sub
End Class
