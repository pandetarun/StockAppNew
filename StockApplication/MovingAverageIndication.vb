Imports FirebirdSql.Data.FirebirdClient
Imports System.Collections.Generic

Public Class MovingAverageIndication
    Public Function GetIndicationIntraDaySMA(tmpStockCode As String) As Boolean

        StockAppLogger.Log("GetIndicationIntraDaySMA Start", "MovingAverageIndication")
        Dim ds As FbDataReader = Nothing
        Dim whereClause As String
        Dim orderClause As String
        Dim configuredPreferredSMAPeriods As String
        Dim tmpPreferredSMAPeriods As List(Of Integer) = Nothing
        Dim lastLowerSMA As Double = 0
        Dim previousLowerSMA As Double = 0
        Dim lastHigherSMA As Double = 0
        Dim previousHigherSMA As Double = 0

        Try
            'Get Preferred SMA for indication
            ds = DBFunctions.getDataFromTable("STOCKWISEPERIODS", " PREFINTRADAYSMAPERIODS ", " stockname='" & tmpStockCode & "'")
            If ds IsNot Nothing And ds.Read() Then
                configuredPreferredSMAPeriods = ds.GetValue(ds.GetOrdinal("PREFINTRADAYSMAPERIODS"))
                tmpPreferredSMAPeriods = New List(Of Integer)(configuredPreferredSMAPeriods.Split(","))
            End If
            ds.Close()
            'Get Lower SMA
            whereClause = "TRADEDDATE='" & Today & "' and STOCKNAME = '" & tmpStockCode & "' and period = " & tmpPreferredSMAPeriods.Item(0)
            orderClause = "lstupdatetime desc"
            ds = DBFunctions.getDataFromTable("INTRADAYSNEMOVINGAVERAGES", " SMA ", whereClause, orderClause)
            If ds.Read() Then
                lastLowerSMA = Double.Parse(ds.GetValue(ds.GetOrdinal("SMA")))
                ds.Read()
                previousLowerSMA = Double.Parse(ds.GetValue(ds.GetOrdinal("SMA")))
            End If
            ds.Close()

            'Get Highr SMA
            whereClause = "TRADEDDATE='" & Today & "' and STOCKNAME = '" & tmpStockCode & "' and period = " & tmpPreferredSMAPeriods.Item(1)
            orderClause = "lstupdatetime desc"
            ds = DBFunctions.getDataFromTable("INTRADAYSNEMOVINGAVERAGES", " SMA ", whereClause, orderClause)
            If ds.Read() Then
                lastHigherSMA = Double.Parse(ds.GetValue(ds.GetOrdinal("SMA")))
                ds.Read()
                previousHigherSMA = Double.Parse(ds.GetValue(ds.GetOrdinal("SMA")))
            End If
            ds.Close()

            'Determine indication

            Return True
        Catch exc As Exception
            StockAppLogger.LogError("GetIndicationIntraDaySMA Error Occurred in indication calculation = ", exc, "MovingAverageIndication")
            Return False
        End Try
        StockAppLogger.Log("GetIndicationIntraDaySMA End", "MovingAverageIndication")
        Return False
    End Function

    Private Function DetermineIndiction(lastLowerSMA As Double, previousLowerSMA As Double, lastHigherSMA As Double, previousHigherSMA As Double) As Boolean

        If previousLowerSMA < previousHigherSMA And lastLowerSMA < lastHigherSMA Then
            'No indication
            'check if already any indication store and if yes then mark that indication false
            'return false
        End If
        If previousLowerSMA < previousHigherSMA And lastLowerSMA > lastHigherSMA Then
            'uptrend started store indication in DB
            'return true
        End If

        If previousLowerSMA - previousHigherSMA > 0 And ((previousLowerSMA - previousHigherSMA) < (lastLowerSMA - lastHigherSMA)) Then
            'uptrend getting stronger
            'store the info
        ElseIf previousLowerSMA - previousHigherSMA > 0 And ((previousLowerSMA - previousHigherSMA) > (lastLowerSMA - lastHigherSMA)) Then
            'uptrend getting weaker
            'store the info
        End If

        If previousLowerSMA > previousHigherSMA And lastLowerSMA < lastHigherSMA Then
            'downtrend started store indication in DB
            'return false
        End If

        Return False
    End Function

End Class
