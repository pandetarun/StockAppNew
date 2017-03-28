Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class MovingAverageCalculation
    Dim threeSampleMinSMA, fiveSampleMinSMA, tenSampleMinSMA, fifteenSampleSMA, twentySampleMinSMA, FiftySampleMinSMA, simpleMA As Double
    Dim threeSampleMinEMA, fiveSampleMinEMA, tenSampleMinEMA, fifteenSampleEMA, twentySampleMinEMA, FiftySampleMinEMA As Double
    Dim MATime, MAStock As String
    Dim MADate As Date
    Dim counter As Integer

    Public Function CalculateAndStoreIntraDayMA() As Boolean
        Dim tmpStockList As List(Of String) = New List(Of String)
        GetStockListAndCalculateIntraDayMA()
        Return False
    End Function

    Public Function CalculateAndStoreDayMA() As Boolean
        Dim tmpStockList As List(Of String) = New List(Of String)
        GetStockListAndCalculateDailyMA()
        Return False
    End Function

    Private Function GetStockListAndCalculateIntraDayMA() As Boolean
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
                    If HourlyMACalculation(tmpStockCode) Then
                        InsertMAtoDB()
                    Else
                        StockAppLogger.Log("GetStockList HourlyMACalcultaion failed for Stock = " & tmpStockCode & " at time = " & MATime, "MovingAverageCalculation")
                    End If
                End If
            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting stocklist from DB = ", exc, "MovingAverageCalculation")
            Return False
        End Try
        StockAppLogger.Log("GetStockList End", "MovingAverageCalculation")
        Return True
    End Function

    Private Function HourlyMACalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim whereClause, whereClause1 As String
        Dim orderClause, orderClause1 As String
        Dim closingPrice As Double
        Dim laststoredEMA As Double
        Dim recordPresentInTable As Boolean

        StockAppLogger.Log("HourlyMACalculation Start", "MovingAverageCalculation")

        counter = 0
        simpleMA = 0
        threeSampleMinSMA = 0
        fiveSampleMinSMA = 0
        tenSampleMinSMA = 0
        fifteenSampleSMA = 0
        twentySampleMinSMA = 0
        FiftySampleMinSMA = 0

        threeSampleMinEMA = 0
        fiveSampleMinEMA = 0
        tenSampleMinEMA = 0
        fifteenSampleEMA = 0
        twentySampleMinEMA = 0
        FiftySampleMinEMA = 0
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
                End If
                simpleMA = simpleMA + Double.Parse(ds.GetValue(ds.GetOrdinal("lastClosingPrice")))
                If counter = 3 Then
                    threeSampleMinSMA = simpleMA / 3
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("THREEEMA")))
                        threeSampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        threeSampleMinEMA = threeSampleMinSMA
                    End If
                ElseIf counter = 5 Then
                    fiveSampleMinSMA = simpleMA / 5
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIVEEMA")))
                        fiveSampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        fiveSampleMinEMA = fiveSampleMinSMA
                    End If
                ElseIf counter = 10 Then
                    tenSampleMinSMA = simpleMA / 10
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("TENEMA")))
                        tenSampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        tenSampleMinEMA = tenSampleMinSMA
                    End If
                ElseIf counter = 15 Then
                    fifteenSampleSMA = simpleMA / 15
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIFTEENEMA")))
                        fifteenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        fifteenSampleEMA = fifteenSampleSMA
                    End If
                ElseIf counter = 20 Then
                    twentySampleMinSMA = simpleMA / 20
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("TWENTYEMA")))
                        twentySampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        twentySampleMinEMA = twentySampleMinSMA
                    End If
                ElseIf counter = 50 Then
                    FiftySampleMinSMA = simpleMA / 50
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIFTYEMA")))
                        FiftySampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        FiftySampleMinEMA = FiftySampleMinSMA
                    End If
                    Exit While
                End If
            End While
            If counter = 1 Then
                threeSampleMinSMA = simpleMA
                fiveSampleMinSMA = simpleMA
                tenSampleMinSMA = simpleMA
                fifteenSampleSMA = simpleMA
                twentySampleMinSMA = simpleMA
                FiftySampleMinSMA = simpleMA
                threeSampleMinEMA = simpleMA
                fiveSampleMinEMA = simpleMA
                tenSampleMinEMA = simpleMA
                fifteenSampleEMA = simpleMA
                twentySampleMinEMA = simpleMA
                FiftySampleMinEMA = simpleMA
            End If
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting hourly stock data from DB = ", exc, "MovingAverageCalculation")
            Return False
        End Try
        StockAppLogger.Log("HourlyMACalculation End", "MovingAverageCalculation")
        Return True
    End Function

    Private Function InsertMAtoDB() As Boolean

        Dim insertStatement As String
        Dim insetValues As String
        Dim sqlStatement As String

        StockAppLogger.Log("InsertMAtoDB Start", "MovingAverageCalculation")
        insertStatement = "INSERT INTO INTRADAYMOVINGAVERAGES (TRADEDDATE, LASTUPDATETIME, STOCK_NAME"
        insetValues = "VALUES ('" & MADate & "','" & MATime & "', '" & MAStock & "'"
        If threeSampleMinSMA <> 0 Then
            insertStatement = insertStatement & ", THREEMA"
            insetValues = insetValues & ", " & threeSampleMinSMA
        End If
        If fiveSampleMinSMA <> 0 Then
            insertStatement = insertStatement & ", FIVEMA"
            insetValues = insetValues & ", " & fiveSampleMinSMA
        End If
        If tenSampleMinSMA <> 0 Then
            insertStatement = insertStatement & ", TENMA"
            insetValues = insetValues & ", " & tenSampleMinSMA
        End If
        If fifteenSampleSMA <> 0 Then
            insertStatement = insertStatement & ", FIFTEENMA"
            insetValues = insetValues & ", " & fifteenSampleSMA
        End If
        If twentySampleMinSMA <> 0 Then
            insertStatement = insertStatement & ", TWENTYMA"
            insetValues = insetValues & ", " & twentySampleMinSMA
        End If
        If FiftySampleMinSMA <> 0 Then
            insertStatement = insertStatement & ", FIFTYMA "
            insetValues = insetValues & ", " & FiftySampleMinSMA
        End If
        If threeSampleMinEMA <> 0 Then
            insertStatement = insertStatement & ", THREEEMA"
            insetValues = insetValues & ", " & threeSampleMinEMA
        End If
        If fiveSampleMinEMA <> 0 Then
            insertStatement = insertStatement & ", FIVEEMA"
            insetValues = insetValues & ", " & fiveSampleMinEMA
        End If
        If tenSampleMinEMA <> 0 Then
            insertStatement = insertStatement & ", TENEMA"
            insetValues = insetValues & ", " & tenSampleMinEMA
        End If
        If fifteenSampleEMA <> 0 Then
            insertStatement = insertStatement & ", FIFTEENEMA"
            insetValues = insetValues & ", " & fifteenSampleEMA
        End If
        If twentySampleMinEMA <> 0 Then
            insertStatement = insertStatement & ", TWENTYEMA"
            insetValues = insetValues & ", " & twentySampleMinEMA
        End If
        If FiftySampleMinEMA <> 0 Then
            insertStatement = insertStatement & ", FIFTYEMA "
            insetValues = insetValues & ", " & FiftySampleMinEMA & " "
        End If
        insertStatement = insertStatement & ") "
        insetValues = insetValues & FiftySampleMinEMA & ");"

        sqlStatement = insertStatement & insetValues
        DBFunctions.ExecuteSQLStmt(sqlStatement)
        StockAppLogger.Log("InsertMAtoDB End", "MovingAverageCalculation")
        Return True
    End Function

    Private Function GetStockListAndCalculateDailyMA() As Boolean
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
                    If DailyMACalculation(tmpStockCode) Then
                        InsertMAtoDB()
                    Else
                        StockAppLogger.Log("GetStockList HourlyMACalcultaion failed for Stock = " & tmpStockCode & " at time = " & MATime, "MovingAverageCalculation")
                    End If
                End If
            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting stocklist from DB = ", exc, "MovingAverageCalculation")
            Return False
        End Try
        StockAppLogger.Log("GetStockList End", "MovingAverageCalculation")
        Return True
    End Function

    Private Function DailyMACalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim whereClause, whereClause1 As String
        Dim orderClause, orderClause1 As String
        Dim closingPrice As Double
        Dim laststoredEMA As Double
        Dim recordPresentInTable As Boolean

        StockAppLogger.Log("DailyMACalculation Start", "MovingAverageCalculation")

        dummy

        counter = 0
        simpleMA = 0
        threeSampleMinSMA = 0
        fiveSampleMinSMA = 0
        tenSampleMinSMA = 0
        fifteenSampleSMA = 0
        twentySampleMinSMA = 0
        FiftySampleMinSMA = 0

        threeSampleMinEMA = 0
        fiveSampleMinEMA = 0
        tenSampleMinEMA = 0
        fifteenSampleEMA = 0
        twentySampleMinEMA = 0
        FiftySampleMinEMA = 0
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
                End If
                simpleMA = simpleMA + Double.Parse(ds.GetValue(ds.GetOrdinal("lastClosingPrice")))
                If counter = 3 Then
                    threeSampleMinSMA = simpleMA / 3
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("THREEEMA")))
                        threeSampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        threeSampleMinEMA = threeSampleMinSMA
                    End If
                ElseIf counter = 5 Then
                    fiveSampleMinSMA = simpleMA / 5
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIVEEMA")))
                        fiveSampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        fiveSampleMinEMA = fiveSampleMinSMA
                    End If
                ElseIf counter = 10 Then
                    tenSampleMinSMA = simpleMA / 10
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("TENEMA")))
                        tenSampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        tenSampleMinEMA = tenSampleMinSMA
                    End If
                ElseIf counter = 15 Then
                    fifteenSampleSMA = simpleMA / 15
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIFTEENEMA")))
                        fifteenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        fifteenSampleEMA = fifteenSampleSMA
                    End If
                ElseIf counter = 20 Then
                    twentySampleMinSMA = simpleMA / 20
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("TWENTYEMA")))
                        twentySampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        twentySampleMinEMA = twentySampleMinSMA
                    End If
                ElseIf counter = 50 Then
                    FiftySampleMinSMA = simpleMA / 50
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIFTYEMA")))
                        FiftySampleMinEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        FiftySampleMinEMA = FiftySampleMinSMA
                    End If
                    Exit While
                End If
            End While
            If counter = 1 Then
                threeSampleMinSMA = simpleMA
                fiveSampleMinSMA = simpleMA
                tenSampleMinSMA = simpleMA
                fifteenSampleSMA = simpleMA
                twentySampleMinSMA = simpleMA
                FiftySampleMinSMA = simpleMA
                threeSampleMinEMA = simpleMA
                fiveSampleMinEMA = simpleMA
                tenSampleMinEMA = simpleMA
                fifteenSampleEMA = simpleMA
                twentySampleMinEMA = simpleMA
                FiftySampleMinEMA = simpleMA
            End If
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting hourly stock data from DB = ", exc, "MovingAverageCalculation")
            Return False
        End Try
        StockAppLogger.Log("HourlyMACalculation End", "MovingAverageCalculation")
        Return True
    End Function

End Class
