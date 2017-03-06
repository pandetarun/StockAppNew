Public Class TestForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tmpNSEindices As NSEindices
        tmpNSEindices = New NSEindices()
        tmpNSEindices.getIndicesListAndStore()
    End Sub
End Class