Module MainModule

    Public Sub main(ByVal parameter As String)
        If parameter = "IndicesData" Then
            Dim tmpNSEindices As NSEindices
            tmpNSEindices = New NSEindices()
            tmpNSEindices.getIndicesListAndStore()
        ElseIf parameter = "IndicesDetails" Then
            Dim tmpNSEIndicesDetails As NSEIndicesDetails
            tmpNSEIndicesDetails = New NSEIndicesDetails()
            tmpNSEIndicesDetails.getIndicesDetailsAndStore()
        ElseIf parameter = "HourlyData" Then
            Dim tmpHourlyStockQuote As HourlyStockQuote
            tmpHourlyStockQuote = New HourlyStockQuote()
            tmpHourlyStockQuote.GetAndStoreHourlyData()
        End If
    End Sub
End Module
