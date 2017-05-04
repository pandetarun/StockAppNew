Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class MovingAverageIndication
    Public Function GetIndicationIntraDaySMA(tmpStockCode As String) As Boolean

        StockAppLogger.Log("GetIndicationIntraDaySMA Start", "MovingAverageIndication")
        Dim ds As FbDataReader = Nothing
        Dim whereClause As String
        Dim orderClause As String
        Dim configuredPreferredSMAPeriods As String
        Dim tmpPreferredSMAPeriods As List(Of String) = Nothing
        Dim lastLowerSMA As Double = 0
        Dim previousLowerSMA As Double = 0
        Dim lastHigherSMA As Double = 0
        Dim previousHigherSMA As Double = 0
        Dim lastupdatetime As String = Nothing
        Dim indicaton As Boolean = False
        Try
            'Get Preferred SMA for indication
            ds = DBFunctions.getDataFromTable("STOCKWISEPERIODS", " PREFINTRADAYSMAPERIODS ", " stockname='" & tmpStockCode & "'")
            If ds IsNot Nothing And ds.Read() Then
                configuredPreferredSMAPeriods = ds.GetValue(ds.GetOrdinal("PREFINTRADAYSMAPERIODS"))
                tmpPreferredSMAPeriods = New List(Of String)(configuredPreferredSMAPeriods.Split(","))
            End If
            ds.Close()
            'Get Lower SMA
            whereClause = "TRADEDDATE='" & Today & "' and STOCKNAME = '" & tmpStockCode & "' and period = " & tmpPreferredSMAPeriods.Item(0)
            orderClause = "lastupdatetime desc"
            ds = DBFunctions.getDataFromTable("INTRADAYSNEMOVINGAVERAGES", " SMA, LASTUPDATETIME ", whereClause, orderClause)
            If ds.Read() Then
                lastupdatetime = ds.GetValue(ds.GetOrdinal("LASTUPDATETIME"))
                lastLowerSMA = Double.Parse(ds.GetValue(ds.GetOrdinal("SMA")))
                ds.Read()
                previousLowerSMA = Double.Parse(ds.GetValue(ds.GetOrdinal("SMA")))
            End If
            ds.Close()

            'Get Highr SMA
            whereClause = "TRADEDDATE='" & Today & "' and STOCKNAME = '" & tmpStockCode & "' and period = " & tmpPreferredSMAPeriods.Item(1)
            orderClause = "lastupdatetime desc"
            ds = DBFunctions.getDataFromTable("INTRADAYSNEMOVINGAVERAGES", " SMA ", whereClause, orderClause)
            If ds.Read() Then
                lastHigherSMA = Double.Parse(ds.GetValue(ds.GetOrdinal("SMA")))
                ds.Read()
                previousHigherSMA = Double.Parse(ds.GetValue(ds.GetOrdinal("SMA")))
            End If
            ds.Close()

            'Determine indication
            indicaton = DetermineIndiction(tmpStockCode, lastupdatetime, lastLowerSMA, previousLowerSMA, lastHigherSMA, previousHigherSMA)
            StockAppLogger.LogInfo("GetIndicationIntraDaySMA Indiation returned = " & indicaton, "MovingAverageIndication")
            Return indicaton
        Catch exc As Exception
            StockAppLogger.LogError("GetIndicationIntraDaySMA Error Occurred in indication calculation = ", exc, "MovingAverageIndication")
            Return False
        End Try
        StockAppLogger.Log("GetIndicationIntraDaySMA End", "MovingAverageIndication")
        Return False
    End Function

    Private Function DetermineIndiction(stockName As String, lastupdatetime As String, lastLowerSMA As Double, previousLowerSMA As Double, lastHigherSMA As Double, previousHigherSMA As Double) As Boolean

        StockAppLogger.Log("DetermineIndiction Start", "MovingAverageIndication")
        Dim whereClause As String
        Dim orderClause As String
        Dim uptrendtimer As Integer = 0
        Dim uptrend As Boolean = False
        Dim ssql As String = Nothing
        Dim ds As FbDataReader = Nothing

        'Get previously stored indication
        whereClause = "INDICATIONDATE='" & Today & "' and STOCKNAME = '" & stockName & "'"
        orderClause = "INDICTATIONTIME desc"
        ds = DBFunctions.getDataFromTable("INTRADAYSMAINDICATION", " UPTRENDINDICATIONTIMER, STRONGUPTREND ", whereClause, orderClause)
        If ds.Read() Then
            uptrendtimer = Integer.Parse(ds.GetValue(ds.GetOrdinal("UPTRENDINDICATIONTIMER")))
            'ds.Read()
            uptrend = ds.GetValue(ds.GetOrdinal("STRONGUPTREND")
        End If
        ds.Close()

        If previousLowerSMA < previousHigherSMA And lastLowerSMA > lastHigherSMA Then
            'uptrend started store indication in DB
            'return true
            'If uptrend Then
            '    DBFunctions.ExecuteSQLStmt("update INTRADAYSMAINDICATION set UPTRENDINDICATIONTIMER=0, INDICATIONDATE='" & Today & "', INDICTATIONTIME='" & lastupdatetime & "' where STOCKNAME ='" & stockName & "' and INDICATIONDATE='" & Today & "'")
            'Else
            ssql = "INSERT INTO INTRADAYSMAINDICATION (STOCKNAME, INDICATIONDATE, INDICTATIONTIME, UPTRENDINDICATIONTIMER) VALUES('" & stockName & "','" & Today & "','" & lastupdatetime & "', 0);"
            DBFunctions.ExecuteSQLStmt(ssql)
            'End If
        End If

        If previousLowerSMA - previousHigherSMA > 0 And ((previousLowerSMA - previousHigherSMA) < (lastLowerSMA - lastHigherSMA)) Then
            'uptrend getting stronger
            'store the info
            If uptrend & uptrendtimer >= 3 Then
                Return True
            Else
                DBFunctions.ExecuteSQLStmt("update INTRADAYSMAINDICATION set UPTRENDINDICATIONTIMER=" & (uptrendtimer + 1) & ", INDICTATIONTIME='" & lastupdatetime & "' where STOCKNAME ='" & stockName & "' and INDICATIONDATE='" & Today & "'")
            End If
        ElseIf previousLowerSMA - previousHigherSMA > 0 And ((previousLowerSMA - previousHigherSMA) > (lastLowerSMA - lastHigherSMA)) Then
            'uptrend getting weaker
            'store the info
            If uptrend Or uptrendtimer >= 0 Then
                DBFunctions.ExecuteSQLStmt("update INTRADAYSMAINDICATION set STRONGUPTREND = 'false', UPTRENDINDICATIONTIMER= - 1, INDICTATIONTIME='" & lastupdatetime & "' where STOCKNAME ='" & stockName & "' and INDICATIONDATE='" & Today & "'")
            End If
            Return False
        End If
        If previousLowerSMA > previousHigherSMA And lastLowerSMA < lastHigherSMA Then
            'downtrend started store indication in DB
            Return False
        End If

        If previousLowerSMA < previousHigherSMA And lastLowerSMA < lastHigherSMA Then
            'No indication
            'check if already any indication store and if yes then mark that indication false
            If uptrend Or uptrendtimer >= 0 Then
                DBFunctions.ExecuteSQLStmt("update INTRADAYSMAINDICATION set STRONGUPTREND = 'false', UPTRENDINDICATIONTIMER= - 1, INDICTATIONTIME='" & lastupdatetime & "' where STOCKNAME ='" & stockName & "' and INDICATIONDATE='" & Today & "'")
            End If
            Return False
        End If
        StockAppLogger.Log("DetermineIndiction End", "MovingAverageIndication")
        Return False
    End Function
End Class
