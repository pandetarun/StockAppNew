Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class MovingAverageCalculation
    Public Function CalculateAndStoreMA() As Boolean
        Dim tmpStockList As List(Of String) = New List(Of String)

        tmpStockList = GetStockList()
        For Each tmpStock In tmpStockList

        Next
        Return False
    End Function

    Private Function GetStockList() As List(Of String)
        Dim tmpStockList As List(Of String) = New List(Of String)
        Dim tmpStockCode As String
        Dim ds As FbDataReader = Nothing

        StockAppLogger.Log("GetStockList Start", "MovingAverageCalculation")
        Try
            ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING")
            While ds.Read()
                tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
                If Not tmpStockList.Contains(tmpStockCode) Then
                    tmpStockList.Add(tmpStockCode)
                End If
            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting stocklist from DB = ", exc, "MovingAverageCalculation")
        End Try
        StockAppLogger.Log("GetStockList End", "MovingAverageCalculation")
        Return tmpStockList
    End Function

    Private Function HourlyMACalculation()

        Dim ds As FbDataReader = Nothing

        StockAppLogger.Log("HourlyMACalculation Start", "MovingAverageCalculation")
        Try
            ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING", "lastClosingPrice", "TradedDate='" & Today & "'")
            While ds.Read()

            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting hourly stock data from DB = ", exc, "MovingAverageCalculation")
        End Try
        StockAppLogger.Log("HourlyMACalculation End", "MovingAverageCalculation")


    End Function
End Class
