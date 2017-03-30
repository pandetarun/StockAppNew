Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class MovingAverage
    Dim lastTradedPrice, threeSampleSMA, fiveSampleSMA, tenSampleSMA, fourteenSampleSMA, twentySampleSMA, FiftySampleSMA, TwoHundredMA, simpleMA As Double
    Dim threeSampleEMA, fiveSampleEMA, tenSampleEMA, fourteenSampleEMA, twentySampleEMA, FiftySampleEMA, TwoHundredEMA As Double

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
                    If IntraDayMACalculation(tmpStockCode) Then
                        InsertIntraDayMAtoDB()
                    Else
                        StockAppLogger.Log("GetStockListAndCalculateIntraDayMA IntraDayMACalcultaion failed for Stock = " & tmpStockCode & " at time = " & MATime, "MovingAverage")
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
        Dim laststoredEMA As Double
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
                    threeSampleSMA = simpleMA / 3
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("THREEEMA")))
                        threeSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        threeSampleEMA = threeSampleSMA
                    End If
                ElseIf counter = 5 Then
                    fiveSampleSMA = simpleMA / 5
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIVEEMA")))
                        fiveSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        fiveSampleEMA = fiveSampleSMA
                    End If
                ElseIf counter = 10 Then
                    tenSampleSMA = simpleMA / 10
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("TENEMA")))
                        tenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        tenSampleEMA = tenSampleSMA
                    End If
                ElseIf counter = 14 Then
                    fourteenSampleSMA = simpleMA / 14
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FOURTEENEMA")))
                        fourteenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        fourteenSampleEMA = fourteenSampleSMA
                    End If
                ElseIf counter = 20 Then
                    twentySampleSMA = simpleMA / 20
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("TWENTYEMA")))
                        twentySampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        twentySampleEMA = twentySampleSMA
                    End If
                ElseIf counter = 50 Then
                    FiftySampleSMA = simpleMA / 50
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIFTYEMA")))
                        FiftySampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
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
            fireQuery = True
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
                    If DailyMACalculation(tmpStockCode) Then
                        InsertDailyMAtoDB()
                    Else
                        StockAppLogger.Log("GetStockList HourlyMACalcultaion failed for Stock = " & tmpStockCode & " at time = " & MATime, "MovingAverage")
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
        Dim laststoredEMA As Double
        Dim recordPresentInTable As Boolean

        StockAppLogger.Log("DailyMACalculation Start", "MovingAverage")
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
            whereClause = "STOCKNAME = '" & tmpStockCode & "' and TRADEDDATE > '" & DateTime.Now.AddDays(-200) & "'"
            orderClause = "TRADEDDATE desc"

            MADate = Today
            'orderClause1 = "lastupdatetime"
            ds = DBFunctions.getDataFromTable("DAILYSTOCKDATA", " LAST_TRADED_PRICE", whereClause, orderClause)
            ds1 = DBFunctions.getDataFromTable("DAILYMOVINGAVERAGES", " * ", whereClause1, orderClause1)
            recordPresentInTable = ds1.Read()
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
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("THREEEMA")))
                        threeSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        threeSampleEMA = threeSampleSMA
                    End If
                ElseIf counter = 5 Then
                    fiveSampleSMA = simpleMA / 5
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIVEEMA")))
                        fiveSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        fiveSampleEMA = fiveSampleSMA
                    End If
                ElseIf counter = 10 Then
                    tenSampleSMA = simpleMA / 10
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("TENEMA")))
                        tenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        tenSampleEMA = tenSampleSMA
                    End If
                ElseIf counter = 14 Then
                    fourteenSampleSMA = simpleMA / 14
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FOURTEENEMA")))
                        fourteenSampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        fourteenSampleEMA = fourteenSampleSMA
                    End If
                ElseIf counter = 20 Then
                    twentySampleSMA = simpleMA / 20
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("TWENTYEMA")))
                        twentySampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        twentySampleEMA = twentySampleSMA
                    End If
                ElseIf counter = 50 Then
                    FiftySampleSMA = simpleMA / 50
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("FIFTYEMA")))
                        FiftySampleEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
                    Else
                        FiftySampleEMA = FiftySampleSMA
                    End If
                ElseIf counter = 200 Then
                    TwoHundredMA = simpleMA / 200
                    If recordPresentInTable Then
                        laststoredEMA = Double.Parse(ds1.GetValue(ds.GetOrdinal("TWOHUNDREDEMA")))
                        TwoHundredEMA = (2 / (counter + 1)) * (closingPrice - laststoredEMA) + laststoredEMA
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
            fireQuery = True
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

End Class
