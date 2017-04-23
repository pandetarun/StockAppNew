Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic
Imports FirebirdSql.Data.Isql

Public Class TMPBollingerBand
    Dim lastTradedPrice, BBUper, BBLower, simpleMA As Double
    Dim PeriodBandwidth As Double
    Dim MATime, MAStock As String
    Dim MADate As Date
    Dim counter As Integer
    Dim insertSQLforallStocks As String = ""

    Public Function CalculateAndStoreIntradayBollingerBands() As Boolean
        Dim tmpStockList As List(Of String) = New List(Of String)
        Dim tmpStockCode As String
        Dim ds As FbDataReader = Nothing

        StockAppLogger.Log("CalculateAndStoreIntradayBollingerBands Start", "BollingerBands")
        Try
            ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING")
            While ds.Read()
                tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
                If Not tmpStockList.Contains(tmpStockCode) Then
                    tmpStockList.Add(tmpStockCode)
                    If IntraDayBBCalculation(tmpStockCode) Then
                        StockAppLogger.LogInfo("CalculateAndStoreIntradayBollingerBands BollingerBands stored for Stock = " & tmpStockCode & " at time = " & MATime, "BollingerBands")
                    Else
                        StockAppLogger.LogInfo("CalculateAndStoreIntradayBollingerBands GetStockListAndCalculateBollingerBands BollingerBands failed for Stock = " & tmpStockCode & " at time = " & MATime, "BollingerBands")
                    End If
                End If
            End While
            InsertIntraDayBBtoDB()
            'DBFunctions.CloseSQLConnection()
        Catch exc As Exception
            StockAppLogger.LogError(" CalculateAndStoreIntradayBollingerBandsError Occurred in getting stocklist from DB = ", exc, "BollingerBands")
            Return False
        End Try
        StockAppLogger.Log("CalculateAndStoreIntradayBollingerBands End", "BollingerBands")
        Return True
    End Function

    Public Function IntraDayBBCalculation(tmpStockCode As String) As Boolean

        Dim ds As FbDataReader = Nothing
        Dim ds1 As FbDataReader = Nothing
        Dim whereClause As String
        Dim orderClause As String
        Dim closingPrice As Double
        Dim tradedPrice, totalTradedPrice As Double
        Dim perioddeviation As Double
        Dim firstPeriodData As List(Of Double) = New List(Of Double)
        Dim secondPeriodData As List(Of Double) = New List(Of Double)
        Dim thirdPeriodData As List(Of Double) = New List(Of Double)
        Dim configuredBBPeriods As String
        Dim tmpBBPeriods As List(Of String)
        Dim tmpPeriodData As List(Of Double)
        Dim tmpsql As String
        Dim totalRecords As Integer = 0
        Dim insertStatement, insertValues As String
        Dim firstperiodSMA, secondperiodSMA, thirdperiodSMA As Double
        Dim tmpdata As Double
        Dim firstperioddeviation, secondperioddeviation, thirdperioddeviation As Double

        StockAppLogger.LogInfo("IntraDayBBCalculation Start", "BollingerBands")
        firstperioddeviation = 0
        secondperioddeviation = 0
        thirdperioddeviation = 0
        tmpdata = 0
        tmpBBPeriods = Nothing
        firstperiodSMA = 0
        secondperiodSMA = 0
        thirdperiodSMA = 0
        lastTradedPrice = 0
        counter = 1
        totalTradedPrice = 0
        simpleMA = 0
        BBUper = 0
        BBLower = 0
        MAStock = tmpStockCode
        Try
            whereClause = "LASTUPDATEDATE='" & Today & "' and companycode = '" & tmpStockCode & "'"
            orderClause = "lastupdatetime desc"
            MADate = Today
            ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " count(lastClosingPrice) totalRows", whereClause)
            If ds.Read() Then
                totalRecords = ds.GetValue(ds.GetOrdinal("totalRows"))
            End If
            ds.Close()
            ds = DBFunctions.getDataFromTable("STOCKHOURLYDATA", " lastClosingPrice, lastupdatetime", whereClause, orderClause)
            ds1 = DBFunctions.getDataFromTable("STOCKWISEPERIODS", " INTRADAYBBPERIOD", "stockname = '" & tmpStockCode & "'")
            If ds1.Read() Then
                configuredBBPeriods = ds1.GetValue(ds1.GetOrdinal("INTRADAYBBPERIOD"))
                tmpBBPeriods = New List(Of String)(configuredBBPeriods.Split(","))
            End If
            ds1.Close()
            If totalRecords >= tmpBBPeriods.Item(0) Then
                insertStatement = " INSERT INTO INTRADAYBOLLINGERBANDS (TRADEDDATE, STOCKNAME, LASTUPDATETIME, PERIOD, CLOSINGPRICE, SMA, UPPERBAND, LOWERBAND, BANDWIDTH) "

                tmpsql = "select sumoffourteentable.sumoffourteen f14, sumoftwentytable.sumoftwenty f20, sumoftwentySixtable.sumoftwentySix f26 from (SELECT SUM(lastclosingprice) sumoffourteen FROM (SELECT first " & tmpBBPeriods.Item(0) & " lastclosingprice FROM STOCKHOURLYDATA WHERE tradeddate='" & Today & "' and companycode='" & tmpStockCode & "' order by LASTUPDATETIME desc)) as sumoffourteentable, (SELECT SUM(lastclosingprice) sumoftwenty FROM (select first " & tmpBBPeriods.Item(1) & " lastclosingprice FROM STOCKHOURLYDATA WHERE tradeddate='" & Today & "' and companycode='" & tmpStockCode & "' order by LASTUPDATETIME desc)) as sumoftwentytable, (SELECT SUM(lastclosingprice) sumoftwentySix FROM (select first " & tmpBBPeriods.Item(2) & " lastclosingprice FROM STOCKHOURLYDATA WHERE tradeddate='" & Today & "' and companycode='" & tmpStockCode & "' order by LASTUPDATETIME desc)) as sumoftwentySixtable;"

                StockAppLogger.Log("IntraDayBBCalculation query Start", "BollingerBands")
                ds1 = DBFunctions.ExecuteSQLStmtandReturnResult(tmpsql)
                If ds1.Read() Then
                    firstperiodSMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("f14"))) / tmpBBPeriods.Item(0)
                    secondperiodSMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("f20"))) / tmpBBPeriods.Item(1)
                    thirdperiodSMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("f26"))) / tmpBBPeriods.Item(2)
                End If
                ds1.Close()
                StockAppLogger.Log("IntraDayBBCalculation query Done", "BollingerBands")
                'If totalRecords >= tmpBBPeriods.Item(1) Then
                '        ds1 = DBFunctions.ExecuteSQLStmtandReturnResult("select sum(lastClosingPrice) lastClosingPrice from (Select First " & tmpBBPeriods.Item(1) & "  lastClosingPrice from STOCKHOURLYDATA where TRADEDDATE='" & Today & "' and companycode = '" & tmpStockCode & "' order by LASTUPDATETIME desc) As T")
                '    If ds1.Read() Then
                '        secondperiodSMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("lastClosingPrice"))) / tmpBBPeriods.Item(0)
                '    End If
                '    ds1.Close()
                'End If

                'If totalRecords >= tmpBBPeriods.Item(2) Then
                '    ds1 = DBFunctions.ExecuteSQLStmtandReturnResult("select sum(lastClosingPrice) lastClosingPrice from (Select First " & tmpBBPeriods.Item(2) & "  lastClosingPrice from STOCKHOURLYDATA where TRADEDDATE='" & Today & "' and companycode = '" & tmpStockCode & "' order by LASTUPDATETIME desc) As T")
                '    If ds1.Read() Then
                '        thirdperiodSMA = Double.Parse(ds1.GetValue(ds1.GetOrdinal("lastClosingPrice"))) / tmpBBPeriods.Item(0)
                '    End If
                '    ds1.Close()
                'End If

                While ds.Read()

                    If counter <= tmpBBPeriods.Item(tmpBBPeriods.Count() - 1) Then
                        If counter = 1 Then
                            MATime = ds.GetValue(ds.GetOrdinal("lastupdatetime"))
                            closingPrice = ds.GetValue(ds.GetOrdinal("lastClosingPrice"))
                            lastTradedPrice = closingPrice
                        End If
                        tradedPrice = Double.Parse(ds.GetValue(ds.GetOrdinal("lastClosingPrice")))
                        totalTradedPrice = totalTradedPrice + tradedPrice
                        tmpdata = tradedPrice - firstperiodSMA
                        'firstPeriodData.Add(tmpdata * tmpdata)
                        firstperioddeviation = firstperioddeviation + (tmpdata * tmpdata)
                        If secondperiodSMA <> 0 Then
                            tmpdata = tradedPrice - secondperiodSMA
                            secondperioddeviation = secondperioddeviation + (tmpdata * tmpdata)
                        End If
                        If thirdperiodSMA <> 0 Then
                            tmpdata = tradedPrice - thirdperiodSMA
                            thirdperioddeviation = thirdperioddeviation + (tmpdata * tmpdata)
                        End If
                        If tmpBBPeriods.Contains(counter) Then
                            If counter = tmpBBPeriods.Item(0) Then
                                perioddeviation = Math.Sqrt(firstperioddeviation)
                                BBLower = firstperiodSMA - 2 * perioddeviation
                                BBUper = firstperiodSMA + 2 * perioddeviation
                                PeriodBandwidth = BBUper - BBLower
                                insertValues = "VALUES ('" & MADate & "', '" & MAStock & "', '" & MATime & "', " & counter & ", " & lastTradedPrice & ", " & firstperiodSMA & ", " & BBUper & ", " & BBLower & ", " & PeriodBandwidth & ");"
                            ElseIf counter = tmpBBPeriods.Item(1) Then
                                perioddeviation = Math.Sqrt(secondperioddeviation)
                                BBLower = secondperiodSMA - 2 * perioddeviation
                                BBUper = secondperiodSMA + 2 * perioddeviation
                                PeriodBandwidth = BBUper - BBLower
                                insertValues = "VALUES ('" & MADate & "', '" & MAStock & "', '" & MATime & "', " & counter & ", " & lastTradedPrice & ", " & secondperiodSMA & ", " & BBUper & ", " & BBLower & ", " & PeriodBandwidth & ");"
                            Else
                                perioddeviation = Math.Sqrt(thirdperioddeviation)
                                BBLower = thirdperiodSMA - 2 * perioddeviation
                                BBUper = thirdperiodSMA + 2 * perioddeviation
                                PeriodBandwidth = BBUper - BBLower
                                insertValues = "VALUES ('" & MADate & "', '" & MAStock & "', '" & MATime & "', " & counter & ", " & lastTradedPrice & ", " & thirdperiodSMA & ", " & BBUper & ", " & BBLower & ", " & PeriodBandwidth & ");"
                            End If
                            insertSQLforallStocks = insertSQLforallStocks & insertStatement & insertValues
                            If tmpBBPeriods.IndexOf(counter) = tmpBBPeriods.Count Then
                                Exit While
                            End If
                        End If
                    Else
                        Exit While
                        End If
                        counter = counter + 1
                    End While

                End If
                ds.Close()
        Catch exc As Exception
            StockAppLogger.LogError("IntraDayBBCalculation Error Occurred in calculating intraday Bollinger Band = ", exc, "BollingerBands")
            Return False
        End Try
        StockAppLogger.LogInfo("IntraDayBBCalculation End", "BollingerBands")
        Return True
    End Function

    Private Sub InsertIntraDayBBtoDB()

        Dim conn As New FbConnection
        Dim command As New FbBatchExecution
        Dim ds As FbDataReader = Nothing

        If DBFunctions.OpenSQLConnection() = True Then
            Try
                'Dim script As String

                Dim fbs As FbScript = New FbScript(insertSQLforallStocks)
                fbs.Parse()
                command = New FbBatchExecution(DBFunctions.myConnection)


                command.AppendSqlStatements(fbs)


                command.Execute(True)

                'DBFunctions.CloseSQLConnection()

                StockAppLogger.Log("getDataFromTable End", "DBFunctions")

            Catch ex As Exception
                StockAppLogger.LogError("getDataFromTable Error in getting data for query ", ex, "DBFunctions")

            End Try
        End If
    End Sub

End Class
