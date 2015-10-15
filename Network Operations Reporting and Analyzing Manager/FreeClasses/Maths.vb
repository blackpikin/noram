Public Class Maths
    Public Function MeanValue(ByVal R As Integer, ByVal C As Integer, ByVal E As Integer, ByVal grid As DataGridView) As Double
        Dim avg As Double = 0.0, n As Integer = 0, sum As Double = 0.0, errMean As String = "/"
        Try
            For i As Integer = C To E
                If IsNumeric(grid.Rows(R).Cells(i).Value) Then
                    n = n + 1
                    sum = sum + grid.Rows(R).Cells(i).Value
                End If
            Next
            avg = Decimal.Round(sum / n, 2, MidpointRounding.AwayFromZero)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return avg
    End Function
End Class
