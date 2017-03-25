Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class HourlyStockQuote

    Public CompanyCode As String
    Public LastUpdateTime As String
    Public TradedDate As Date
    Public FreeFloatMarketCapInCrs As Double
    Public bcStartDate As Date
    Public ChangeFromPreviousDay As Double
    Public buyQuantity3 As Double
    Public sellPrice1 As Double
    Public buyQuantity4 As Double
    Public sellPrice2 As Double
    Public priceBand As String
    Public BuyQuantity As Double
    Public deliveryQuantity As Double
    Public buyQuantity2 As Double
    Public sellPrice5 As Double
    Public TradedVolumeShares As Integer
    Public buyQuantity5 As Double
    Public sellPrice3 As Double
    Public sellPrice4 As Double
    Public OpenPrice As Double
    Public Low52 As Double
    Public securityVar As Double
    Public marketType As String
    Public UpperPriceBand As Double
    Public TotalTradedValueInLacs As Double
    Public FaceValue As Double
    Public ndStartDate As Date
    Public PreviousDayClosePrice As Double
    Public varMargin As Double
    Public LastClosingPrice As Double
    Public PercentageChange As Double
    Public adhocMargin As Double
    Public CompanyName As String
    Public VolumeWeightedAveragePrice As Double
    Public SecDate As Date
    Public ISINCode As String
    Public indexVar As Double
    Public LowerPriceBand As Double
    Public totalBuyQuantity As Double
    Public high52 As Double
    Public purpose As String
    Public cm_adj_low_dt As Date
    Public closePrice As Double
    Public isExDateFlag As String
    Public recordDate As Date
    Public cm_adj_high_dt As Date
    Public totalSellQuantity As Double
    Public dayHigh As Double
    Public exDate As Date
    Public sellQuantity5 As Double
    Public bcEndDate As Date
    Public css_status_desc As String
    Public ndEndDate As Date
    Public sellQuantity2 As Double
    Public sellQuantity1 As Double
    Public buyPrice1 As Double
    Public sellQuantity4 As Double
    Public buyPrice2 As Double
    Public sellQuantity3 As Double
    Public applicableMargin As Double
    Public buyPrice4 As Double
    Public buyPrice3 As Double
    Public buyPrice5 As Double
    Public dayLow As Double
    Public deliveryToTradedQuantity As Double
    Public totalTradedVolume As Double
    Public lastUpdateDate As Date

    Public insertStatement As String

    Public Function GetAndStoreHourlyData() As Boolean

        Dim tmpStockList As List(Of String) = New List(Of String)
        Dim tmpStockCode As String
        Dim ds As FbDataReader = Nothing
        Dim rawHourlyStockQuote As String
        Dim tmpHourlyStockQuote As HourlyStockQuote

        StockAppLogger.Log("GetAndStoreHourlyData Start")
        ds = DBFunctions.getDataFromTable("NSE_INDICES_TO_STOCK_MAPPING")
        While ds.Read()
            tmpStockCode = ds.GetValue(ds.GetOrdinal("STOCK_NAME"))
            tmpStockCode = tmpStockCode.Replace("&", "%26")
            If Not tmpStockList.Contains(tmpStockCode) Then
                StockAppLogger.Log("GetAndStoreHourlyData started running for stock = " & tmpStockCode)
                rawHourlyStockQuote = Helper.GetDataFromUrl(My.Settings.TimelyStockQuote & tmpStockCode)
                tmpHourlyStockQuote = CreateObjectFromRawStockData(rawHourlyStockQuote)
                Try
                    StockAppLogger.Log("GetAndStoreHourlyData stock = " & tmpHourlyStockQuote.CompanyCode & " insert statement = " & tmpHourlyStockQuote.insertStatement)
                    DBFunctions.ExecuteSQLStmt(tmpHourlyStockQuote.insertStatement)
                Catch exc As Exception
                    StockAppLogger.LogError("Error Occurred in inserting hourlystockdata = ", exc)
                End Try
                StockAppLogger.Log("GetAndStoreHourlyData end running for stock = " & tmpStockCode)
                tmpStockList.Add(tmpStockCode)
            End If
        End While
        DBFunctions.CloseSQLConnection()

        StockAppLogger.Log("GetAndStoreHourlyData End")
        Return True
    End Function

    Public Function CreateObjectFromRawStockData(ByVal rawStockQuote As String) As HourlyStockQuote

        Dim myDelims As String() = New String() {","""}
        Dim quoteLines() As String = rawStockQuote.Substring(rawStockQuote.IndexOf("lastUpdateTime")).Split(myDelims, StringSplitOptions.None)

        StockAppLogger.Log("CreateObjectFromRawStockData Start")
        Dim hourlyQuoteTemp As HourlyStockQuote = New HourlyStockQuote()
        insertStatement = CreateInsertStatement()
        'Get first line for last update date time
        myDelims = New String() {""":"}

        hourlyQuoteTemp = AssignValuestoObject(hourlyQuoteTemp, quoteLines)

        StockAppLogger.Log("CreateObjectFromRawStockData End")
        Return hourlyQuoteTemp
    End Function

    Public Function AssignValuestoObject(ByVal hourlyQuoteTemp As HourlyStockQuote, ByVal quoteLines As String()) As HourlyStockQuote
        Dim quoteItemTag As String
        Dim quoteItemValue As String
        Dim myDelims As String() = New String() {""":"}
        Dim lastQuoteLine() As String
        Dim insertColumns As String
        Dim insertValues As String
        Dim lastUpdateDateTime As String

        StockAppLogger.Log("AssignValuestoObject Start")
        insertColumns = "INSERT INTO STOCKHOURLYDATA ("
        insertValues = "values ("
        'Get rest of the parameters
        Try
            For Each quoteLine As String In quoteLines
                If Not quoteLine.Contains("futLink") And Not quoteLine.Contains("otherSeries") And Not quoteLine.Contains("optLink") Then
                    quoteItemValue = quoteLine.Split(myDelims, StringSplitOptions.None)(1)
                    quoteItemTag = quoteLine.Split(myDelims, StringSplitOptions.None)(0)
                    If Not quoteItemValue.Contains("""-") Then
                        If quoteItemTag.Contains("symbol") Then
                            hourlyQuoteTemp.CompanyCode = quoteItemValue.Replace("""", "")
                            insertColumns = insertColumns + "COMPANYCODE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.CompanyCode.Replace("'", "''") + "',"
                        ElseIf quoteItemTag.Contains("lastUpdateTime") Then
                            lastUpdateDateTime = quoteItemValue.Replace("""", "")
                            hourlyQuoteTemp.lastUpdateDate = Date.Parse(lastUpdateDateTime.Split(" ")(0))
                            hourlyQuoteTemp.LastUpdateTime = lastUpdateDateTime.Split(" ")(1)
                            insertColumns = insertColumns + "LASTUPDATETIME,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.LastUpdateTime + "',"
                            insertColumns = insertColumns + "LASTUPDATEDATE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.lastUpdateDate + "',"
                        ElseIf quoteItemTag.Contains("tradedDate") Then
                            hourlyQuoteTemp.TradedDate = Date.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "TRADEDDATE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.TradedDate + "',"
                        ElseIf quoteItemTag.Contains("cm_ffm") Then
                            hourlyQuoteTemp.FreeFloatMarketCapInCrs = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "FREEFLOATMARKETCAPINCRS,"
                            insertValues = insertValues + hourlyQuoteTemp.FreeFloatMarketCapInCrs.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("bcStartDate") Then
                            hourlyQuoteTemp.bcStartDate = Date.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BCSTARTDATE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.bcStartDate + "',"
                        ElseIf quoteItemTag.Contains("change") Then
                            hourlyQuoteTemp.ChangeFromPreviousDay = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "CHANGEFROMPREVIOUSDAY,"
                            insertValues = insertValues + hourlyQuoteTemp.ChangeFromPreviousDay.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("buyQuantity3") Then
                            hourlyQuoteTemp.buyQuantity3 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BUYQUANTITY3,"
                            insertValues = insertValues + hourlyQuoteTemp.buyQuantity3.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("buyQuantity4") Then
                            hourlyQuoteTemp.buyQuantity4 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BUYQUANTITY4,"
                            insertValues = insertValues + hourlyQuoteTemp.buyQuantity4.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("sellPrice2") Then
                            hourlyQuoteTemp.sellPrice2 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SELLPRICE2,"
                            insertValues = insertValues + hourlyQuoteTemp.sellPrice2.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("priceBand") Then
                            hourlyQuoteTemp.priceBand = quoteItemValue.Replace("""", "")
                            insertColumns = insertColumns + "PRICEBAND,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.priceBand.Replace("'", "''") + "',"
                        ElseIf quoteItemTag.Contains("buyQuantity1") Then
                            hourlyQuoteTemp.BuyQuantity = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BUYQUANTITY,"
                            insertValues = insertValues + hourlyQuoteTemp.BuyQuantity.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("deliveryQuantity") Then
                            hourlyQuoteTemp.deliveryQuantity = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "DELIVERYQUANTITY,"
                            insertValues = insertValues + hourlyQuoteTemp.deliveryQuantity.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("buyQuantity2") Then
                            hourlyQuoteTemp.buyQuantity2 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BUYQUANTITY2,"
                            insertValues = insertValues + hourlyQuoteTemp.buyQuantity2.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("sellPrice5") Then
                            hourlyQuoteTemp.sellPrice5 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SELLPRICE5,"
                            insertValues = insertValues + hourlyQuoteTemp.sellPrice5.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("quantityTraded") Then
                            hourlyQuoteTemp.deliveryToTradedQuantity = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "TRADEDVOLUMESHARES,"
                            insertValues = insertValues + hourlyQuoteTemp.deliveryToTradedQuantity.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("buyQuantity5") Then
                            hourlyQuoteTemp.buyQuantity5 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BUYQUANTITY5,"
                            insertValues = insertValues + hourlyQuoteTemp.buyQuantity5.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("sellPrice3") Then
                            hourlyQuoteTemp.sellPrice3 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SELLPRICE3,"
                            insertValues = insertValues + hourlyQuoteTemp.sellPrice3.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("sellPrice4") Then
                            hourlyQuoteTemp.sellPrice4 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SELLPRICE4,"
                            insertValues = insertValues + hourlyQuoteTemp.sellPrice4.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("open") Then
                            hourlyQuoteTemp.OpenPrice = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "OPENPRICE,"
                            insertValues = insertValues + hourlyQuoteTemp.OpenPrice.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("low52") Then
                            hourlyQuoteTemp.Low52 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "LOW52,"
                            insertValues = insertValues + hourlyQuoteTemp.Low52.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("securityVar") Then
                            hourlyQuoteTemp.securityVar = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SECURITYVAR,"
                            insertValues = insertValues + hourlyQuoteTemp.securityVar.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("marketType") Then
                            hourlyQuoteTemp.marketType = quoteItemValue.Replace("""", "")
                            insertColumns = insertColumns + "MARKETTYPE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.marketType.Replace("'", "''") + "',"
                        ElseIf quoteItemTag.Contains("pricebandupper") Then
                            hourlyQuoteTemp.UpperPriceBand = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "UPPERPRICEBAND,"
                            insertValues = insertValues + hourlyQuoteTemp.UpperPriceBand.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("totalTradedValue") Then
                            hourlyQuoteTemp.TotalTradedValueInLacs = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "TOTALTRADEDVALUEINLACS,"
                            insertValues = insertValues + hourlyQuoteTemp.TotalTradedValueInLacs.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("faceValue") Then
                            hourlyQuoteTemp.FaceValue = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "FACEVALUE,"
                            insertValues = insertValues + hourlyQuoteTemp.FaceValue.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("ndStartDate") Then
                            hourlyQuoteTemp.ndStartDate = Date.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "NDSTARTDATE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.ndStartDate.ToString("R") + "',"
                        ElseIf quoteItemTag.Contains("previousClose") Then
                            hourlyQuoteTemp.PreviousDayClosePrice = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "PREVIOUSDAYCLOSEPRICE,"
                            insertValues = insertValues + hourlyQuoteTemp.PreviousDayClosePrice.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("varMargin") Then
                            hourlyQuoteTemp.varMargin = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "VARMARGIN,"
                            insertValues = insertValues + hourlyQuoteTemp.varMargin.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("lastPrice") Then
                            hourlyQuoteTemp.LastClosingPrice = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "LASTCLOSINGPRICE,"
                            insertValues = insertValues + hourlyQuoteTemp.LastClosingPrice.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("pChange") Then
                            hourlyQuoteTemp.PercentageChange = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "PERCENTAGECHANGE,"
                            insertValues = insertValues + hourlyQuoteTemp.PercentageChange.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("adhocMargin") Then
                            '   If Not quoteItemValue.Contains("""-") Then
                            hourlyQuoteTemp.adhocMargin = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "ADHOCMARGIN,"
                            insertValues = insertValues + hourlyQuoteTemp.adhocMargin.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("companyName") Then
                            hourlyQuoteTemp.CompanyName = quoteItemValue.Replace("""", "")
                            insertColumns = insertColumns + "COMPANYNAME,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.CompanyName.Replace("'", "''") + "',"
                        ElseIf quoteItemTag.Contains("averagePrice") Then
                            hourlyQuoteTemp.VolumeWeightedAveragePrice = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "VOLUMEWEIGHTEDAVERAGEPRICE,"
                            insertValues = insertValues + hourlyQuoteTemp.VolumeWeightedAveragePrice.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("secDate") Then
                            hourlyQuoteTemp.SecDate = Date.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SECDATE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.SecDate + "',"
                        ElseIf quoteItemTag.Contains("isinCode") Then
                            hourlyQuoteTemp.ISINCode = quoteItemValue.Replace("""", "")
                            insertColumns = insertColumns + "ISINCODE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.ISINCode.Replace("'", "''") + "',"
                        ElseIf quoteItemTag.Contains("indexVar") Then
                            hourlyQuoteTemp.indexVar = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "INDEXVAR,"
                            insertValues = insertValues + hourlyQuoteTemp.indexVar.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("pricebandlower") Then
                            hourlyQuoteTemp.LowerPriceBand = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "LOWERPRICEBAND,"
                            insertValues = insertValues + hourlyQuoteTemp.LowerPriceBand.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("totalBuyQuantity") Then
                            hourlyQuoteTemp.totalBuyQuantity = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "TOTALBUYQUANTITY,"
                            insertValues = insertValues + hourlyQuoteTemp.totalBuyQuantity.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("high52") Then
                            hourlyQuoteTemp.high52 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "HIGH52,"
                            insertValues = insertValues + hourlyQuoteTemp.high52.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("purpose") Then
                            hourlyQuoteTemp.purpose = quoteItemValue.Replace("""", "")
                            insertColumns = insertColumns + "PURPOSE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.purpose.Replace("'", "''") + "',"
                        ElseIf quoteItemTag.Contains("cm_adj_low_dt") Then
                            hourlyQuoteTemp.cm_adj_low_dt = Date.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "CM_ADJ_LOW_DT,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.cm_adj_low_dt + "',"
                        ElseIf quoteItemTag.Contains("closePrice") Then
                            hourlyQuoteTemp.closePrice = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "CLOSEPRICE,"
                            insertValues = insertValues + hourlyQuoteTemp.closePrice.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("isExDateFlag") Then
                            hourlyQuoteTemp.isExDateFlag = quoteItemValue.Replace("""", "")
                            insertColumns = insertColumns + "ISEXDATEFLAG,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.isExDateFlag + "',"
                        ElseIf quoteItemTag.Contains("recordDate") Then
                            hourlyQuoteTemp.recordDate = Date.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "RECORDDATE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.recordDate + "',"
                        ElseIf quoteItemTag.Contains("cm_adj_high_dt") Then
                            hourlyQuoteTemp.cm_adj_high_dt = Date.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "CM_ADJ_HIGH_DT,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.cm_adj_high_dt + "',"
                        ElseIf quoteItemTag.Contains("totalSellQuantity") Then
                            hourlyQuoteTemp.totalSellQuantity = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "TOTALSELLQUANTITY,"
                            insertValues = insertValues + hourlyQuoteTemp.totalSellQuantity.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("dayHigh") Then
                            hourlyQuoteTemp.dayHigh = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "DAYHIGH,"
                            insertValues = insertValues + hourlyQuoteTemp.dayHigh.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("exDate") Then
                            hourlyQuoteTemp.exDate = Date.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "EXDATE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.exDate + "',"
                        ElseIf quoteItemTag.Contains("sellQuantity5") Then
                            hourlyQuoteTemp.sellQuantity5 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SELLQUANTITY5,"
                            insertValues = insertValues + hourlyQuoteTemp.sellQuantity5.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("bcEndDate") Then
                            hourlyQuoteTemp.bcEndDate = Date.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BCENDDATE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.bcEndDate + "',"
                        ElseIf quoteItemTag.Contains("css_status_desc") Then
                            hourlyQuoteTemp.css_status_desc = quoteItemValue.Replace("""", "")
                            insertColumns = insertColumns + "CSS_STATUS_DESC,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.css_status_desc + "',"
                        ElseIf quoteItemTag.Contains("ndEndDate") Then
                            hourlyQuoteTemp.ndEndDate = Date.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "NDENDDATE,"
                            insertValues = insertValues + "'" + hourlyQuoteTemp.ndEndDate + "',"
                        ElseIf quoteItemTag.Contains("sellQuantity2") Then
                            hourlyQuoteTemp.sellQuantity2 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SELLQUANTITY2,"
                            insertValues = insertValues + hourlyQuoteTemp.sellQuantity2.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("sellQuantity1") Then
                            hourlyQuoteTemp.sellQuantity1 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SELLQUANTITY1,"
                            insertValues = insertValues + hourlyQuoteTemp.sellQuantity1.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("buyPrice1") Then
                            hourlyQuoteTemp.buyPrice1 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BUYPRICE1,"
                            insertValues = insertValues + hourlyQuoteTemp.buyPrice1.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("sellQuantity4") Then
                            hourlyQuoteTemp.sellQuantity4 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SELLQUANTITY4,"
                            insertValues = insertValues + hourlyQuoteTemp.sellQuantity4.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("buyPrice2") Then
                            hourlyQuoteTemp.buyPrice2 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BUYPRICE2,"
                            insertValues = insertValues + hourlyQuoteTemp.buyPrice2.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("sellQuantity3") Then
                            hourlyQuoteTemp.sellQuantity3 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SELLQUANTITY3,"
                            insertValues = insertValues + hourlyQuoteTemp.sellQuantity3.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("applicableMargin") Then
                            hourlyQuoteTemp.applicableMargin = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "APPLICABLEMARGIN,"
                            insertValues = insertValues + hourlyQuoteTemp.applicableMargin.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("buyPrice4") Then
                            hourlyQuoteTemp.buyPrice4 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BUYPRICE4,"
                            insertValues = insertValues + hourlyQuoteTemp.buyPrice4.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("buyPrice3") Then
                            hourlyQuoteTemp.buyPrice3 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BUYPRICE3,"
                            insertValues = insertValues + hourlyQuoteTemp.buyPrice3.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("buyPrice5") Then
                            hourlyQuoteTemp.buyPrice5 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "BUYPRICE5,"
                            insertValues = insertValues + hourlyQuoteTemp.buyPrice5.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("dayLow") Then
                            hourlyQuoteTemp.dayLow = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "DAYLOW,"
                            insertValues = insertValues + hourlyQuoteTemp.dayLow.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("deliveryToTradedQuantity") Then
                            hourlyQuoteTemp.deliveryToTradedQuantity = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "DELIVERYTOTRADEDQUANTITY,"
                            insertValues = insertValues + hourlyQuoteTemp.deliveryToTradedQuantity.ToString("R") + ","
                        ElseIf quoteItemTag.Contains("totalTradedVolume") Then
                            myDelims = New String() {"""}]"}
                            lastQuoteLine = quoteItemValue.Split(myDelims, StringSplitOptions.None)
                            hourlyQuoteTemp.totalTradedVolume = Double.Parse(lastQuoteLine(0).Replace("""", ""))
                            insertColumns = insertColumns + "TOTALTRADEDVOLUME"
                            insertValues = insertValues + hourlyQuoteTemp.totalTradedVolume.ToString("R")
                        ElseIf quoteItemTag.Contains("sellPrice1") Then
                            hourlyQuoteTemp.sellPrice1 = Double.Parse(quoteItemValue.Replace("""", ""))
                            insertColumns = insertColumns + "SELLPRICE1,"
                            insertValues = insertValues + hourlyQuoteTemp.sellPrice1.ToString("R") + ","
                        End If
                    End If
                End If
            Next
        Catch exc As Exception
            StockAppLogger.LogError("Error Occurred in creating hourlystock object = ", exc)
        End Try
        insertColumns = insertColumns + ") "
        insertValues = insertValues + ");"
        insertStatement = insertColumns + insertValues
        hourlyQuoteTemp.insertStatement = insertStatement
        StockAppLogger.Log("AssignValuestoObject End")
        Return hourlyQuoteTemp
    End Function

    Private Function CreateInsertStatement() As String
        Dim sqlStatement As String
        sqlStatement = "INSERT INTO STOCKHOURLYDATA (LASTUPDATETIME, TRADEDDATE, FREEFLOATMARKETCAPINCRS, BCSTARTDATE, CHANGEFROMPREVIOUSDAY, BUYQUANTITY3, SELLPRICE1, BUYQUANTITY4, SELLPRICE2, PRICEBAND, BUYQUANTITY, DELIVERYQUANTITY, BUYQUANTITY2, SELLPRICE5, TRADEDVOLUMESHARES, BUYQUANTITY5, SELLPRICE3, SELLPRICE4, OPENPRICE, LOW52, SECURITYVAR, MARKETTYPE, UPPERPRICEBAND, TOTALTRADEDVALUEINLACS, FACEVALUE, NDSTARTDATE, PREVIOUSDAYCLOSEPRICE, COMPANYCODE, VARMARGIN, LASTCLOSINGPRICE, PERCENTAGECHANGE, ADHOCMARGIN, COMPANYNAME, VOLUMEWEIGHTEDAVERAGEPRICE, SECDATE, ISINCODE, INDEXVAR, LOWERPRICEBAND, TOTALBUYQUANTITY, HIGH52, PURPOSE, CM_ADJ_LOW_DT, CLOSEPRICE, ISEXDATEFLAG, RECORDDATE, CM_ADJ_HIGH_DT, TOTALSELLQUANTITY, DAYHIGH, EXDATE, SELLQUANTITY5, BCENDDATE, CSS_STATUS_DESC, NDENDDATE, SELLQUANTITY2, SELLQUANTITY1, BUYPRICE1, SELLQUANTITY4, BUYPRICE2, SELLQUANTITY3, APPLICABLEMARGIN, BUYPRICE4, BUYPRICE3, BUYPRICE5, DAYLOW, DELIVERYTOTRADEDQUANTITY, TOTALTRADEDVOLUME) values("
        Return sqlStatement
    End Function

End Class
