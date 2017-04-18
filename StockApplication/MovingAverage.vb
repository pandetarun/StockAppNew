Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class MovingAverage
    Dim lastTradedPrice, threeSampleSMA, fiveSampleSMA, tenSampleSMA, fourteenSampleSMA, twentySampleSMA, FiftySampleSMA, TwoHundredMA, simpleMA As Double
    Dim eMA, threeSampleEMA, fiveSampleEMA, tenSampleEMA, fourteenSampleEMA, twentySampleEMA, FiftySampleEMA, TwoHundredEMA As Double

    Dim MATime, MAStock As String
    Dim MADate As Date
    Dim counter As Integer

    Public Function CalculateAndStoreIntraDayMA() As Boolean
        'Dim tmpStockList As List(Of String) = New List(Of String)
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

        StockAppLogger.Log("GetStockListAndCalculateIntraDayMA Start", "MovingAverage")
        Try
            ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING")
            While ds.Read()
                tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
                If Not tmpStockList.Contains(tmpStockCode) Then
                    tmpStockList.Add(tmpStockCode)
                    If IntraDaySNEMACalculation(tmpStockCode) Then
                        StockAppLogger.Log("GetStockListAndCalculateIntraDayMA IntraDaySNEMACalculation successfull for Stock = " & tmpStockCode & " at time = " & MATime, "MovingAverage")
                    Else
                        StockAppLogger.LogInfo("GetStockListAndCalculateIntraDayMA IntraDaySNEMACalculation failed for Stock = " & tmpStockCode & " at time = " & MATime, "MovingAverage")
                    End If
                End If
            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting stocklist from DB = ", exc, "MovingAverage")
            Return False
        End Try
        StockAppLogger.Log("GetStockListAndCalculateIntraDayMA End", "MovingAverage")
        Return True
    End Function

    Private Function IntraDayMACalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim whereClause, whereClause1 As String
        Dim orderClause, orderClause1 As String
        Dim closingPrice As Double
        Dim laststoredThreeEMA, laststoredFiveEMA, laststoredTenEMA, laststoredFourteenEMA, laststoredTwentyEMA, laststoredFiftyEMA As Double
        Dim recordPresentInTable As Boolean

        StockAppLogger.Log("IntraDayMACalculation Start", "MovingAverage")

        lastTradedPrice = 0
        counter = 0
        simpleMA = 0
        threeSampleSMA = 0
        fiveSampleSMA = 0
        tenSampleSMA = 0
        fourteenSampleSMA = 0
        twentySampleSMA = 0
        FiftySampleSMA = 0
        laststoredThreeEMA = 0
        laststoredFiveEMA = 0
        laststoredTenEMA = 0
        laststoredFourteenEMA = 0
        laststoredTwentyEMA = 0
        laststoredFiftyEMA = 0
        threeSampleEMA = 0
        fiveSampleEMA = 0
        tenSampleEMA = 0
        fourteenSampleEMA = 0
        twentySampleEMA = 0
        FiftySampleEMA = 0
        MAStock = tmpStockCode

        Try
            whereClause1 = "TRADEDDATE='" & Today & "' and STOCK_NAME = '" & tmpStockCode & "'"
            orderClause1 = "LASTUPDATETIME desc"
            whereClause = "LASTUPDATEDATE='" & Today & "' and companycode = '" & tmpStockCode & "'"
            orderClause = "lastupdatetime desc"

            MADate = Today
            'orderClause1 = "lastupdatetime"
            ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " lastClosingPrice, lastupdatetime", whereClause, orderClause)
            ds1 = DBFunctions.getDataFromTable("INTRADAYMOVINGAVERAGES", " * ", whereClause1, orderClause1)

            While ds1.Read()
                recordPresentInTable = True
                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("THREEEMA"))) Then
                    laststoredThreeEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("THREEEMA")))
                End If
                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("FIVEEMA"))) Then
                    laststoredFiveEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("FIVEEMA")))
                End If

                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("TENEMA"))) Then
                    laststoredTenEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("TENEMA")))
                End If
                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("FOURTEENEMA"))) Then
                    laststoredFourteenEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("FOURTEENEMA")))
                End If
                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("TWENTYEMA"))) Then
                    laststoredTwentyEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("TWENTYEMA")))
                End If
                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("FIFTYEMA"))) Then
                    laststoredFiftyEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("FIFTYEMA")))
                End If
            End While

            While ds.Read()
                counter = counter + 1
                If counter = 1 Then
                    MATime = ds.GetValue(ds.GetOrdinal("lastupdatetime"))
                    closingPrice = ds.GetValue(ds.GetOrdinal("lastClosingPrice"))
                    lastTradedPrice = closingPrice
                End If
                simpleMA = simpleMA + Double.Parse(ds.GetValue(ds.GetOrdinal("lastClosingPrice")))
                If counter = 3 Then
                    threeSampleSMA = simpleMA / 3
                    If recordPresentInTable And laststoredThreeEMA > 0 Then
                        threeSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredThreeEMA) + laststoredThreeEMA
                    Else
                        threeSampleEMA = threeSampleSMA
                    End If
                ElseIf counter = 5 Then
                    fiveSampleSMA = simpleMA / 5
                    If recordPresentInTable And laststoredFiveEMA > 0 Then
                        fiveSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredFiveEMA) + laststoredFiveEMA
                    Else
                        fiveSampleEMA = fiveSampleSMA
                    End If
                ElseIf counter = 10 Then
                    tenSampleSMA = simpleMA / 10
                    If recordPresentInTable And laststoredTenEMA > 0 Then
                        tenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredTenEMA) + laststoredTenEMA
                    Else
                        tenSampleEMA = tenSampleSMA
                    End If
                ElseIf counter = 14 Then
                    fourteenSampleSMA = simpleMA / 14
                    If recordPresentInTable And laststoredFourteenEMA > 0 Then
                        fourteenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredFourteenEMA) + laststoredFourteenEMA
                    Else
                        fourteenSampleEMA = fourteenSampleSMA
                    End If
                ElseIf counter = 20 Then
                    twentySampleSMA = simpleMA / 20
                    If recordPresentInTable And laststoredTwentyEMA > 0 Then
                        twentySampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredTwentyEMA) + laststoredTwentyEMA
                    Else
                        twentySampleEMA = twentySampleSMA
                    End If
                ElseIf counter = 50 Then
                    FiftySampleSMA = simpleMA / 50
                    If recordPresentInTable And laststoredFiftyEMA > 0 Then
                        FiftySampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredFiftyEMA) + laststoredFiftyEMA
                    Else
                        FiftySampleEMA = FiftySampleSMA
                    End If
                    Exit While
                End If
            End While
            If counter = 1 Then
                threeSampleSMA = simpleMA
                fiveSampleSMA = simpleMA
                tenSampleSMA = simpleMA
                fourteenSampleSMA = simpleMA
                twentySampleSMA = simpleMA
                FiftySampleSMA = simpleMA
                threeSampleEMA = simpleMA
                fiveSampleEMA = simpleMA
                tenSampleEMA = simpleMA
                fourteenSampleEMA = simpleMA
                twentySampleEMA = simpleMA
                FiftySampleEMA = simpleMA
            End If
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in calculating intraday moving average = ", exc, "MovingAverage")
            Return False
        End Try
        StockAppLogger.Log("IntraDayMACalculation End", "MovingAverage")
        Return True
    End Function

    Public Function IntraDaySNEMACalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim orderClause, whereClause, configuredSMAPeriods, configuredEMAPeriods As String
        Dim closingPrice As Double
        Dim laststoredEMA, tmpSimpleMA As Double
        Dim tmpMAPeriods, tmpEMAPeriods As List(Of String)

        StockAppLogger.Log("IntraDaySNEMACalculation Start", "MovingAverage")

        tmpMAPeriods = Nothing
        tmpEMAPeriods = Nothing
        lastTradedPrice = 0
        counter = 0
        simpleMA = 0
        eMA = 0
        tmpSimpleMA = 0
        laststoredEMA = 0
        MAStock = tmpStockCode
        Try
            whereClause = "LASTUPDATEDATE='" & Today & "' and companycode = '" & tmpStockCode & "'"
            orderClause = "lastupdatetime desc"
            MADate = Today
            ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " lastClosingPrice, lastupdatetime", whereClause, orderClause)
            'Get period details for the stock
            ds1 = DBFunctions.getDataFromTable("STOCKWISEPERIODS", " INTRADAYSMAPERIOD, INTRADAYEMAPERIOD", "stockname = '" & tmpStockCode & "'")
            If ds1.Read() Then
                configuredSMAPeriods = ds1.GetValue(ds1.GetOrdinal("INTRADAYSMAPERIOD"))
                configuredEMAPeriods = ds1.GetValue(ds1.GetOrdinal("INTRADAYEMAPERIOD"))
                tmpMAPeriods = New List(Of String)(configuredSMAPeriods.Split(","))
                tmpEMAPeriods = New List(Of String)(configuredEMAPeriods.Split(","))
            End If
            ds1.Close()
            While ds.Read()
                counter = counter + 1
                If counter = 1 Then
                    MATime = ds.GetValue(ds.GetOrdinal("lastupdatetime"))
                    closingPrice = ds.GetValue(ds.GetOrdinal("lastClosingPrice"))
                    lastTradedPrice = closingPrice
                End If
                tmpSimpleMA = tmpSimpleMA + Double.Parse(ds.GetValue(ds.GetOrdinal("lastClosingPrice")))
                If tmpMAPeriods.Contains(counter) Then
                    simpleMA = tmpSimpleMA / counter
                    CalculateEMAForPeriod(counter, tmpStockCode, closingPrice)
                    If InsertIntraDaySNEMAtoDB(counter) Then
                        StockAppLogger.Log("IntraDaySNEMACalculation period " & counter & " inserted in DB for stock = " & tmpStockCode, "MovingAverage")
                    Else
                        StockAppLogger.LogInfo("IntraDaySNEMACalculation period " & counter & " not inserted in DB  for stock = " & tmpStockCode, "MovingAverage")
                    End If
                    If tmpMAPeriods.IndexOf(counter) = tmpMAPeriods.Count Then
                        Exit While
                    End If
                End If
            End While
            ds.Close()
            If counter = 1 Then
                simpleMA = tmpSimpleMA
                If InsertIntraDaySNEMAtoDB(counter) Then
                    StockAppLogger.Log("IntraDaySNEMACalculation period " & counter & " inserted in DB for stock = " & tmpStockCode, "MovingAverage")
                Else
                    StockAppLogger.LogInfo("IntraDaySNEMACalculation period " & counter & " not inserted in DB  for stock = " & tmpStockCode, "MovingAverage")
                End If
            End If
        Catch exc As Exception
            StockAppLogger.LogError("IntraDaySNEMACalculation Error Occurred in calculating intraday moving average = ", exc, "MovingAverage")
            Return False
        End Try
        StockAppLogger.Log("IntraDaySNEMACalculation End", "MovingAverage")
        Return True
    End Function

    Private Function CalculateEMAForPeriod(ByVal period As Integer, tmpStockCode As String, closingprice As Double) As Boolean
        Dim orderClause, whereClause As String
        Dim ds As FbDataReader = Nothing
        Dim laststoredEMA As Double

        laststoredEMA = 0
        whereClause = "TRADEDDATE='" & Today & "' and STOCKNAME = '" & tmpStockCode & "' and PERIOD = " & counter
        orderClause = "LASTUPDATETIME desc"
        ds = DBFunctions.getDataFromTable("INTRADAYSNEMOVINGAVERAGES", " * ", whereClause, orderClause)
        If ds.Read() Then
            laststoredEMA = Double.Parse(ds.GetValue(ds.GetOrdinal("EMA")))
        End If
        If laststoredEMA <> 0 Then
            eMA = (2 / (counter + 1)) * (closingprice - laststoredEMA) + laststoredEMA
        Else
            eMA = simpleMA
        End If
        Return True
    End Function

    Private Function InsertIntraDayMAtoDB() As Boolean

        Dim insertStatement As String
        Dim insertValues As String
        Dim sqlStatement As String
        Dim fireQuery As Boolean = False

        StockAppLogger.Log("InsertIntraDayMAtoDB Start", "MovingAverage")
        insertStatement = "INSERT INTO INTRADAYMOVINGAVERAGES (TRADEDDATE, LASTUPDATETIME, STOCK_NAME"
        insertValues = "VALUES ('" & MADate & "','" & MATime & "', '" & MAStock & "'"
        If lastTradedPrice <> 0 Then
            insertStatement = insertStatement & ", LASTTRADEDPRICE"
            insertValues = insertValues & ", " & lastTradedPrice
        End If
        If threeSampleSMA <> 0 Then
            insertStatement = insertStatement & ", THREEMA"
            insertValues = insertValues & ", " & threeSampleSMA
            fireQuery = True
        End If
        If fiveSampleSMA <> 0 Then
            insertStatement = insertStatement & ", FIVEMA"
            insertValues = insertValues & ", " & fiveSampleSMA
            fireQuery = True
        End If
        If tenSampleSMA <> 0 Then
            insertStatement = insertStatement & ", TENMA"
            insertValues = insertValues & ", " & tenSampleSMA
            fireQuery = True
        End If
        If fourteenSampleSMA <> 0 Then
            insertStatement = insertStatement & ", FOURTEENMA"
            insertValues = insertValues & ", " & fourteenSampleSMA
            fireQuery = True
        End If
        If twentySampleSMA <> 0 Then
            insertStatement = insertStatement & ", TWENTYMA"
            insertValues = insertValues & ", " & twentySampleSMA
            fireQuery = True
        End If
        If FiftySampleSMA <> 0 Then
            insertStatement = insertStatement & ", FIFTYMA "
            insertValues = insertValues & ", " & FiftySampleSMA
            fireQuery = True
        End If
        If threeSampleEMA <> 0 Then
            insertStatement = insertStatement & ", THREEEMA"
            insertValues = insertValues & ", " & threeSampleEMA
            fireQuery = True
        End If
        If fiveSampleEMA <> 0 Then
            insertStatement = insertStatement & ", FIVEEMA"
            insertValues = insertValues & ", " & fiveSampleEMA
            fireQuery = True
        End If
        If tenSampleEMA <> 0 Then
            insertStatement = insertStatement & ", TENEMA"
            insertValues = insertValues & ", " & tenSampleEMA
            fireQuery = True
        End If
        If fourteenSampleEMA <> 0 Then
            insertStatement = insertStatement & ", FOURTEENEMA"
            insertValues = insertValues & ", " & fourteenSampleEMA
            fireQuery = True
        End If
        If twentySampleEMA <> 0 Then
            insertStatement = insertStatement & ", TWENTYEMA"
            insertValues = insertValues & ", " & twentySampleEMA
            fireQuery = True
        End If
        If FiftySampleEMA <> 0 Then
            insertStatement = insertStatement & ", FIFTYEMA "
            insertValues = insertValues & ", " & FiftySampleEMA & " "
            fireQuery = True
        End If
        insertStatement = insertStatement & ") "
        insertValues = insertValues & ");"
        sqlStatement = insertStatement & insertValues
        If fireQuery Then
            DBFunctions.ExecuteSQLStmt(sqlStatement)
        Else
            StockAppLogger.Log("InsertIntraDayMAtoDB Insert Query not fired for stock = " & MAStock & " at time = " & MATime, "MovingAverage")
        End If
        StockAppLogger.Log("InsertIntraDayMAtoDB End", "MovingAverage")
        Return True
    End Function

    Private Function InsertIntraDaySNEMAtoDB(ByVal period As Integer) As Boolean

        Dim insertStatement As String
        Dim insertValues As String
        Dim sqlStatement As String
        Dim fireQuery As Boolean = False

        StockAppLogger.Log("InsertIntraDaySNEMAtoDB Start", "MovingAverage")
        insertStatement = "INSERT INTO INTRADAYSNEMOVINGAVERAGES (TRADEDDATE, LASTUPDATETIME, STOCKNAME, CLOSINGPRICE, SMA, EMA, PERIOD"
        insertValues = "VALUES ('" & MADate & "','" & MATime & "', '" & MAStock & "'," & lastTradedPrice & ", " & simpleMA & ", " & eMA & ", " & period

        insertStatement = insertStatement & ") "
        insertValues = insertValues & ");"
        sqlStatement = insertStatement & insertValues

        DBFunctions.ExecuteSQLStmt(sqlStatement)
        DBFunctions.CloseSQLConnection()
        StockAppLogger.Log("InsertIntraDaySNEMAtoDB End", "MovingAverage")
        Return True
    End Function

    Private Function GetStockListAndCalculateDailyMA() As Boolean
        Dim tmpStockList As List(Of String) = New List(Of String)
        Dim tmpStockCode As String
        Dim ds As FbDataReader = Nothing

        StockAppLogger.Log("GetStockList Start", "MovingAverage")
        Try
            ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING")
            While ds.Read()
                tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
                If Not tmpStockList.Contains(tmpStockCode) Then
                    tmpStockList.Add(tmpStockCode)
                    If DailySNEMACalculation(tmpStockCode) Then
                        StockAppLogger.Log("GetStockList HourlyMACalcultaion successfull for Stock = " & tmpStockCode & " at Date = " & MADate, "MovingAverage")
                    Else
                        StockAppLogger.LogInfo("GetStockList HourlyMACalcultaion failed for Stock = " & tmpStockCode & " at Date = " & MADate, "MovingAverage")
                    End If
                End If
            End While
            DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in getting stocklist from DB = ", exc, "MovingAverage")
            Return False
        End Try
        StockAppLogger.Log("GetStockList End", "MovingAverage")
        Return True
    End Function

    Private Function DailyMACalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim whereClause, whereClause1 As String
        Dim orderClause, orderClause1 As String
        Dim closingPrice As Double
        Dim laststoredThreeEMA, laststoredFiveEMA, laststoredTenEMA, laststoredFourteenEMA, laststoredTwentyEMA, laststoredFiftyEMA, laststoredTwoHundredEMA As Double
        Dim recordPresentInTable As Boolean

        StockAppLogger.Log("DailyMACalculation Start", "MovingAverage")
        laststoredThreeEMA = 0
        laststoredFiveEMA = 0
        laststoredTenEMA = 0
        laststoredFourteenEMA = 0
        laststoredTwentyEMA = 0
        laststoredFiftyEMA = 0
        laststoredTwoHundredEMA = 0
        lastTradedPrice = 0
        counter = 0
        simpleMA = 0
        threeSampleSMA = 0
        fiveSampleSMA = 0
        tenSampleSMA = 0
        fourteenSampleSMA = 0
        twentySampleSMA = 0
        FiftySampleSMA = 0
        TwoHundredMA = 0
        threeSampleEMA = 0
        fiveSampleEMA = 0
        tenSampleEMA = 0
        fourteenSampleEMA = 0
        twentySampleEMA = 0
        FiftySampleEMA = 0
        TwoHundredEMA = 0
        MAStock = tmpStockCode

        Try
            whereClause1 = "STOCK_NAME = '" & tmpStockCode & "'"
            orderClause1 = "TRADEDDATE desc"
            whereClause = "STOCKNAME = '" & tmpStockCode & "' and TRADEDDATE > '" & Today.AddDays(-200) & "'"
            orderClause = "TRADEDDATE desc"

            MADate = Today
            'orderClause1 = "lastupdatetime"
            ds = DBFunctions.getDataFromTable("DAILYSTOCKDATA", " LAST_TRADED_PRICE", whereClause, orderClause)
            ds1 = DBFunctions.getDataFromTable("DAILYMOVINGAVERAGES", " * ", whereClause1, orderClause1)
            While ds1.Read()
                recordPresentInTable = True
                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("THREEEMA"))) Then
                    laststoredThreeEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("THREEEMA")))
                End If

                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("FIVEEMA"))) Then
                    laststoredFiveEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("FIVEEMA")))
                End If

                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("TENEMA"))) Then
                    laststoredTenEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("TENEMA")))
                End If
                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("FOURTEENEMA"))) Then
                    laststoredFourteenEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("FOURTEENEMA")))
                End If
                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("TWENTYEMA"))) Then
                    laststoredTwentyEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("TWENTYEMA")))
                End If
                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("FIFTYEMA"))) Then
                    laststoredFiftyEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("FIFTYEMA")))
                End If
                If Not IsDBNull(ds1.GetValue(ds1.GetOrdinal("TWOHUNDREDEMA"))) Then
                    laststoredTwoHundredEMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("TWOHUNDREDEMA")))
                End If
            End While
            While ds.Read()
                counter = counter + 1
                If counter = 1 Then
                    'MATime = ds.GetValue(ds.GetOrdinal("lastupdatetime"))
                    closingPrice = ds.GetValue(ds.GetOrdinal("LAST_TRADED_PRICE"))
                    lastTradedPrice = closingPrice
                End If
                simpleMA = simpleMA + Double.Parse(ds.GetValue(ds.GetOrdinal("LAST_TRADED_PRICE")))
                If counter = 3 Then
                    threeSampleSMA = simpleMA / 3
                    If recordPresentInTable And laststoredThreeEMA > 0 Then
                        threeSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredThreeEMA) + laststoredThreeEMA
                    Else
                        threeSampleEMA = threeSampleSMA
                    End If
                ElseIf counter = 5 Then
                    fiveSampleSMA = simpleMA / 5
                    If recordPresentInTable And laststoredFiveEMA > 0 Then
                        fiveSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredFiveEMA) + laststoredFiveEMA
                    Else
                        fiveSampleEMA = fiveSampleSMA
                    End If
                ElseIf counter = 10 Then
                    tenSampleSMA = simpleMA / 10
                    If recordPresentInTable And laststoredTenEMA > 0 Then
                        tenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredTenEMA) + laststoredTenEMA
                    Else
                        tenSampleEMA = tenSampleSMA
                    End If
                ElseIf counter = 14 Then
                    fourteenSampleSMA = simpleMA / 14
                    If recordPresentInTable And laststoredFourteenEMA > 0 Then
                        fourteenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredFourteenEMA) + laststoredFourteenEMA
                    Else
                        fourteenSampleEMA = fourteenSampleSMA
                    End If
                ElseIf counter = 20 Then
                    twentySampleSMA = simpleMA / 20
                    If recordPresentInTable And laststoredTwentyEMA > 0 Then
                        twentySampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredTwentyEMA) + laststoredTwentyEMA
                    Else
                        twentySampleEMA = twentySampleSMA
                    End If
                ElseIf counter = 50 Then
                    FiftySampleSMA = simpleMA / 50
                    If recordPresentInTable And laststoredFiftyEMA > 0 Then
                        FiftySampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredFiftyEMA) + laststoredFiftyEMA
                    Else
                        FiftySampleEMA = FiftySampleSMA
                    End If
                ElseIf counter = 200 Then
                    TwoHundredMA = simpleMA / 200
                    If recordPresentInTable And laststoredTwoHundredEMA > 0 Then
                        TwoHundredEMA = (2 / (counter + 1)) * (closingPrice - laststoredTwoHundredEMA) + laststoredTwoHundredEMA
                    Else
                        TwoHundredEMA = TwoHundredMA
                    End If
                    Exit While
                End If

            End While
            If counter = 1 Then
                threeSampleSMA = simpleMA
                fiveSampleSMA = simpleMA
                tenSampleSMA = simpleMA
                fourteenSampleSMA = simpleMA
                twentySampleSMA = simpleMA
                FiftySampleSMA = simpleMA
                TwoHundredMA = simpleMA
                threeSampleEMA = simpleMA
                fiveSampleEMA = simpleMA
                tenSampleEMA = simpleMA
                fourteenSampleEMA = simpleMA
                twentySampleEMA = simpleMA
                FiftySampleEMA = simpleMA
                TwoHundredEMA = simpleMA
            End If
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in calculating daily MA data = ", exc, "MovingAverage")
            Return False
        End Try
        StockAppLogger.Log("DailyMACalculation End", "MovingAverage")
        Return True
    End Function

    Public Function DailySNEMACalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim orderClause, whereClause, configuredSMAPeriods, configuredEMAPeriods As String
        Dim closingPrice As Double
        Dim laststoredEMA, tmpSimpleMA As Double
        Dim tmpMAPeriods, tmpEMAPeriods As List(Of String)

        StockAppLogger.Log("DailySNEMACalculation Start", "MovingAverage")

        tmpMAPeriods = Nothing
        tmpEMAPeriods = Nothing
        lastTradedPrice = 0
        counter = 0
        simpleMA = 0
        eMA = 0
        tmpSimpleMA = 0
        laststoredEMA = 0
        MAStock = tmpStockCode
        Try
            whereClause = "STOCKNAME = '" & tmpStockCode & "' and TRADEDDATE > '" & Today.AddDays(-200) & "'"
            orderClause = "TRADEDDATE desc"
            MADate = Today
            ds = DBFunctions.getDataFromTable("DAILYSTOCKDATA", " last_traded_price", whereClause, orderClause)
            'Get period details for the stock
            ds1 = DBFunctions.getDataFromTable("STOCKWISEPERIODS", " DAILYSMAPERIOD, DAILYEMAPERIOD", "stockname = '" & tmpStockCode & "'")
            If ds1.Read() Then
                configuredSMAPeriods = ds1.GetValue(ds1.GetOrdinal("DAILYSMAPERIOD"))
                configuredEMAPeriods = ds1.GetValue(ds1.GetOrdinal("DAILYEMAPERIOD"))
                tmpMAPeriods = New List(Of String)(configuredSMAPeriods.Split(","))
                tmpEMAPeriods = New List(Of String)(configuredEMAPeriods.Split(","))
            End If
            ds1.Close()
            While ds.Read()
                counter = counter + 1
                If counter = 1 Then
                    closingPrice = ds.GetValue(ds.GetOrdinal("last_traded_price"))
                    lastTradedPrice = closingPrice
                End If
                tmpSimpleMA = tmpSimpleMA + Double.Parse(ds.GetValue(ds.GetOrdinal("LAST_TRADED_PRICE")))
                If tmpMAPeriods.Contains(counter) Then
                    simpleMA = tmpSimpleMA / counter
                    CalculateDailyEMAForPeriod(counter, tmpStockCode, closingPrice)
                    If InsertDailySNEMAtoDB(counter) Then
                        StockAppLogger.Log("DailySNEMACalculation period " & counter & " inserted in DB for stock = " & tmpStockCode, "MovingAverage")
                    Else
                        StockAppLogger.LogInfo("DailySNEMACalculation period " & counter & " not inserted in DB  for stock = " & tmpStockCode, "MovingAverage")
                    End If
                    If tmpMAPeriods.IndexOf(counter) = tmpMAPeriods.Count Then
                        Exit While
                    End If
                End If
            End While
            ds.Close()
            If counter = 1 Then
                simpleMA = tmpSimpleMA
                If InsertIntraDaySNEMAtoDB(counter) Then
                    StockAppLogger.Log("DailySNEMACalculation period " & counter & " inserted in DB for stock = " & tmpStockCode, "MovingAverage")
                Else
                    StockAppLogger.LogInfo("DailySNEMACalculation period " & counter & " not inserted in DB  for stock = " & tmpStockCode, "MovingAverage")
                End If
            End If
        Catch exc As Exception
            StockAppLogger.LogError("DailySNEMACalculation Error Occurred in calculating intraday moving average = ", exc, "MovingAverage")
            Return False
        End Try
        StockAppLogger.Log("DailySNEMACalculation End", "MovingAverage")
        Return True
    End Function

    Private Function CalculateDailyEMAForPeriod(ByVal period As Integer, tmpStockCode As String, closingprice As Double) As Boolean
        Dim orderClause, whereClause As String
        Dim ds As FbDataReader = Nothing
        Dim laststoredEMA As Double

        laststoredEMA = 0
        whereClause = "STOCKNAME = '" & tmpStockCode & "' and PERIOD = " & counter
        orderClause = "TRADEDDATE desc"
        ds = DBFunctions.getDataFromTable("DAILYSNEMOVINGAVERAGES", " * ", whereClause, orderClause)
        If ds.Read() Then
            laststoredEMA = Double.Parse(ds.GetValue(ds.GetOrdinal("EMA")))
        End If
        If laststoredEMA <> 0 Then
            eMA = (2 / (counter + 1)) * (closingprice - laststoredEMA) + laststoredEMA
        Else
            eMA = simpleMA
        End If
        Return True
    End Function

    Private Function InsertDailyMAtoDB() As Boolean

        Dim insertStatement As String
        Dim insertValues As String
        Dim sqlStatement As String
        Dim fireQuery As Boolean = False

        StockAppLogger.Log("InsertDailyMAtoDB Start", "MovingAverage")
        insertStatement = "INSERT INTO DAILYMOVINGAVERAGES (TRADEDDATE, STOCK_NAME"
        insertValues = "VALUES ('" & MADate & "', '" & MAStock & "'"
        If lastTradedPrice <> 0 Then
            insertStatement = insertStatement & ", LASTTRADEDPRICE"
            insertValues = insertValues & ", " & lastTradedPrice
        End If
        If threeSampleSMA <> 0 Then
            insertStatement = insertStatement & ", THREEMA"
            insertValues = insertValues & ", " & threeSampleSMA
            fireQuery = True
        End If
        If fiveSampleSMA <> 0 Then
            insertStatement = insertStatement & ", FIVEMA"
            insertValues = insertValues & ", " & fiveSampleSMA
            fireQuery = True
        End If
        If tenSampleSMA <> 0 Then
            insertStatement = insertStatement & ", TENMA"
            insertValues = insertValues & ", " & tenSampleSMA
            fireQuery = True
        End If
        If fourteenSampleSMA <> 0 Then
            insertStatement = insertStatement & ", FOURTEENMA"
            insertValues = insertValues & ", " & fourteenSampleSMA
            fireQuery = True
        End If
        If twentySampleSMA <> 0 Then
            insertStatement = insertStatement & ", TWENTYMA"
            insertValues = insertValues & ", " & twentySampleSMA
            fireQuery = True
        End If
        If FiftySampleSMA <> 0 Then
            insertStatement = insertStatement & ", FIFTYMA "
            insertValues = insertValues & ", " & FiftySampleSMA
            fireQuery = True
        End If
        If TwoHundredMA <> 0 Then
            insertStatement = insertStatement & ", TwoHundredMA "
            insertValues = insertValues & ", " & TwoHundredMA
            fireQuery = True
        End If
        If threeSampleEMA <> 0 Then
            insertStatement = insertStatement & ", THREEEMA"
            insertValues = insertValues & ", " & threeSampleEMA
            fireQuery = True
        End If
        If fiveSampleEMA <> 0 Then
            insertStatement = insertStatement & ", FIVEEMA"
            insertValues = insertValues & ", " & fiveSampleEMA
            fireQuery = True
        End If
        If tenSampleEMA <> 0 Then
            insertStatement = insertStatement & ", TENEMA"
            insertValues = insertValues & ", " & tenSampleEMA
            fireQuery = True
        End If
        If fourteenSampleEMA <> 0 Then
            insertStatement = insertStatement & ", FOURTEENEMA"
            insertValues = insertValues & ", " & fourteenSampleEMA
            fireQuery = True
        End If
        If twentySampleEMA <> 0 Then
            insertStatement = insertStatement & ", TWENTYEMA"
            insertValues = insertValues & ", " & twentySampleEMA
            fireQuery = True
        End If
        If FiftySampleEMA <> 0 Then
            insertStatement = insertStatement & ", FIFTYEMA "
            insertValues = insertValues & ", " & FiftySampleEMA
            fireQuery = True
        End If
        If TwoHundredEMA <> 0 Then
            insertStatement = insertStatement & ", TwoHundredEMA "
            insertValues = insertValues & ", " & TwoHundredEMA
            fireQuery = True
        End If
        insertStatement = insertStatement & ") "
        insertValues = insertValues & ");"

        sqlStatement = insertStatement & insertValues
        If fireQuery Then
            DBFunctions.ExecuteSQLStmt(sqlStatement)
        Else
            StockAppLogger.Log("InsertDailyMAtoDB Insert Query not fired for stock = " & MAStock & " at Date = " & MADate, "MovingAverage")
        End If
        StockAppLogger.Log("InsertDailyMAtoDB End", "MovingAverage")
        Return True
    End Function

    Private Function InsertDailySNEMAtoDB(ByVal period As Integer) As Boolean
        Dim insertStatement As String
        Dim insertValues As String
        Dim sqlStatement As String
        Dim fireQuery As Boolean = False

        StockAppLogger.Log("InsertDailySNEMAtoDB Start", "MovingAverage")
        insertStatement = "INSERT INTO DAILYSNEMOVINGAVERAGES (TRADEDDATE, STOCKNAME, CLOSINGPRICE, SMA, EMA, PERIOD"
        insertValues = "VALUES ('" & MADate & "', '" & MAStock & "'," & lastTradedPrice & ", " & simpleMA & ", " & eMA & ", " & period

        insertStatement = insertStatement & ") "
        insertValues = insertValues & ");"
        sqlStatement = insertStatement & insertValues

        DBFunctions.ExecuteSQLStmt(sqlStatement)

        StockAppLogger.Log("InsertDailySNEMAtoDB End", "MovingAverage")
        Return True
    End Function

End Class
