Imports System.Windows.Forms.DataVisualization.Charting

Public Class GraphRSLRTN

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
        RTNRSL.Show()
    End Sub

    Private Sub GraphRSLRTN_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        
    End Sub

    Private Function MeanValue(ByVal R As Integer, ByVal C As Integer, ByVal E As Integer, ByVal grid As DataGridView) As Double
        Dim avg As Double = 0.0, n As Integer = 0, sum As Double = 0.0
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


    Public Sub PlotBadRSL(ByVal numberofDays As Integer)
        Chart1.Series.Clear()
        For Each row As DataGridViewRow In RTNRSL.DataGridView2.SelectedRows
            Try
                If numberofDays = 3 Then
                    With Chart1
                        .Series.Add(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 6, 9, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 10, 13, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 14, 17, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                        .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                        .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                        .ChartAreas("ChartArea1").AxisX.Title = "Days"
                        .ChartAreas("ChartArea1").AxisY.Title = "RSL mean (db)"
                        ' Set scrollbar size
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                        ' Show small scroll buttons only
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                        ' Scrollbars position
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                        ' Change scrollbar colors
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                    End With

                ElseIf numberofDays = 4 Then
                    With Chart1
                        .Series.Add(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 6, 9, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 10, 13, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 14, 17, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 18, 21, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                        .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                        .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                        .ChartAreas("ChartArea1").AxisX.Title = "Days"
                        .ChartAreas("ChartArea1").AxisY.Title = "RSL mean (db)"
                        ' Set scrollbar size
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                        ' Show small scroll buttons only
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                        ' Scrollbars position
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                        ' Change scrollbar colors
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                    End With

                ElseIf numberofDays = 5 Then
                    With Chart1
                        .Series.Add(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 6, 9, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 10, 13, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 14, 17, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 18, 21, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(5, MeanValue(row.Index, 22, 25, RTNRSL.DataGridView2))

                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                        .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                        .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                        .ChartAreas("ChartArea1").AxisX.Title = "Days"
                        .ChartAreas("ChartArea1").AxisY.Title = "RSL mean (db)"
                        ' Set scrollbar size
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                        ' Show small scroll buttons only
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                        ' Scrollbars position
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                        ' Change scrollbar colors
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                    End With
                ElseIf numberofDays = 6 Then
                    With Chart1
                        .Series.Add(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 6, 9, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 10, 13, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 14, 17, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 18, 21, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(5, MeanValue(row.Index, 22, 25, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(6, MeanValue(row.Index, 26, 29, RTNRSL.DataGridView2))

                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                        .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                        .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                        .ChartAreas("ChartArea1").AxisX.Title = "Days"
                        .ChartAreas("ChartArea1").AxisY.Title = "RSL mean (db)"
                        ' Set scrollbar size
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                        ' Show small scroll buttons only
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                        ' Scrollbars position
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                        ' Change scrollbar colors
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                    End With
                ElseIf numberofDays = 7 Then
                    With Chart1
                        .Series.Add(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 6, 9, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 10, 13, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 14, 17, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 18, 21, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(5, MeanValue(row.Index, 22, 25, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(6, MeanValue(row.Index, 26, 29, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(7, MeanValue(row.Index, 30, 33, RTNRSL.DataGridView2))

                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                        .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                        .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                        .ChartAreas("ChartArea1").AxisX.Title = "Days"
                        .ChartAreas("ChartArea1").AxisY.Title = "RSL mean (db)"
                        ' Set scrollbar size
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                        ' Show small scroll buttons only
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                        ' Scrollbars position
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                        ' Change scrollbar colors
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                    End With
                End If
            Catch ex As Exception
                MessageBox.Show("The data contains only non numerical values." & vbLf & "The mean could not be calculated." & vbLf & "Division by zero is impossible!")
            End Try
        Next
    End Sub

    Public Sub PlotDI(ByVal numberOfDays As Integer)
        Chart1.Series.Clear()
        For Each row As DataGridViewRow In RTNRSL.DataGridView2.SelectedRows
            Try
                If numberOfDays = 3 Then
                    With Chart1
                        .Series.Add(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 8, 9, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 12, 13, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 16, 17, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                        .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                        .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                        .ChartAreas("ChartArea1").AxisX.Title = "Days"
                        .ChartAreas("ChartArea1").AxisY.Title = "RSL mean (db)"
                        ' Set scrollbar size
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                        ' Show small scroll buttons only
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                        ' Scrollbars position
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                        ' Change scrollbar colors
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                    End With

                ElseIf numberOfDays = 4 Then
                    With Chart1
                        .Series.Add(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 8, 9, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 12, 13, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 16, 17, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 20, 21, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                        .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                        .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                        .ChartAreas("ChartArea1").AxisX.Title = "Days"
                        .ChartAreas("ChartArea1").AxisY.Title = "RSL mean (db)"
                        ' Set scrollbar size
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                        ' Show small scroll buttons only
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                        ' Scrollbars position
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                        ' Change scrollbar colors
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                    End With

                ElseIf numberOfDays = 5 Then
                    With Chart1
                        .Series.Add(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 8, 9, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 12, 13, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 16, 17, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 20, 21, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(5, MeanValue(row.Index, 24, 25, RTNRSL.DataGridView2))

                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                        .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                        .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                        .ChartAreas("ChartArea1").AxisX.Title = "Days"
                        .ChartAreas("ChartArea1").AxisY.Title = "RSL mean (db)"
                        ' Set scrollbar size
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                        ' Show small scroll buttons only
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                        ' Scrollbars position
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                        ' Change scrollbar colors
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                    End With
                ElseIf numberOfDays = 6 Then
                    With Chart1
                        .Series.Add(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 8, 9, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 12, 13, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 16, 17, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 20, 21, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(5, MeanValue(row.Index, 24, 25, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(6, MeanValue(row.Index, 28, 29, RTNRSL.DataGridView2))

                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                        .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                        .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                        .ChartAreas("ChartArea1").AxisX.Title = "Days"
                        .ChartAreas("ChartArea1").AxisY.Title = "RSL mean (db)"
                        ' Set scrollbar size
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                        ' Show small scroll buttons only
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                        ' Scrollbars position
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                        ' Change scrollbar colors
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                    End With
                ElseIf numberOfDays = 7 Then
                    With Chart1
                        .Series.Add(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 8, 9, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 12, 13, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 16, 17, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 20, 21, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(5, MeanValue(row.Index, 24, 25, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(6, MeanValue(row.Index, 28, 29, RTNRSL.DataGridView2))
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(7, MeanValue(row.Index, 32, 33, RTNRSL.DataGridView2))

                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                        .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                        .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                        .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                        .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                        .ChartAreas("ChartArea1").AxisX.Title = "Days"
                        .ChartAreas("ChartArea1").AxisY.Title = "RSL mean (db)"
                        ' Set scrollbar size
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                        ' Show small scroll buttons only
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                        ' Scrollbars position
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                        ' Change scrollbar colors
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                        .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                        .Series(RTNRSL.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & RTNRSL.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                    End With
                End If
            Catch ex As Exception
                MessageBox.Show("The data contains only non numerical values." & vbLf & "The mean could not be calculated." & vbLf & "Division by zero is impossible!")
            End Try
        Next
    End Sub

    Public Sub PlotBadEthernet(ByVal numberOfDays As Integer)
        Chart1.Series.Clear()
        For Each row As DataGridViewRow In EtheCapacity.DataGridView2.SelectedRows
            'Try
            If numberOfDays = 3 Then
                With Chart1
                    .Series.Add(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value)
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 6, 7, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 8, 9, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 10, 11, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                    .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                    .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                    .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                    .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                    .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                    .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                    .ChartAreas("ChartArea1").AxisX.Title = "Days"
                    .ChartAreas("ChartArea1").AxisY.Title = "Ethernet capacity mean (Mbit/s)"
                    ' Set scrollbar size
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                    ' Show small scroll buttons only
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                    ' Scrollbars position
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                    ' Change scrollbar colors
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                End With

            ElseIf numberOfDays = 4 Then
                With Chart1
                    .Series.Add(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value)
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 6, 7, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 8, 9, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 10, 11, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 12, 13, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                    .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                    .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                    .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                    .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                    .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                    .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                    .ChartAreas("ChartArea1").AxisX.Title = "Days"
                    .ChartAreas("ChartArea1").AxisY.Title = "Ethernet capacity mean (Mbit/s)"
                    ' Set scrollbar size
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                    ' Show small scroll buttons only
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                    ' Scrollbars position
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                    ' Change scrollbar colors
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                End With

            ElseIf numberOfDays = 5 Then
                With Chart1
                    .Series.Add(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value)
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 6, 7, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 8, 9, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 10, 11, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 12, 13, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(5, MeanValue(row.Index, 14, 15, EtheCapacity.DataGridView2))

                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                    .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                    .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                    .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                    .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                    .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                    .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                    .ChartAreas("ChartArea1").AxisX.Title = "Days"
                    .ChartAreas("ChartArea1").AxisY.Title = "Ethernet capacity mean (Mbit/s)"
                    ' Set scrollbar size
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                    ' Show small scroll buttons only
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                    ' Scrollbars position
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                    ' Change scrollbar colors
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                End With
            ElseIf numberOfDays = 6 Then
                With Chart1
                    .Series.Add(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value)
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 6, 7, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 8, 9, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 10, 11, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 12, 13, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(5, MeanValue(row.Index, 14, 15, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(6, MeanValue(row.Index, 16, 17, EtheCapacity.DataGridView2))

                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                    .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                    .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                    .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                    .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                    .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                    .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                    .ChartAreas("ChartArea1").AxisX.Title = "Days"
                    .ChartAreas("ChartArea1").AxisY.Title = "Ethernet capacity mean (Mbit/s)"
                    ' Set scrollbar size
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                    ' Show small scroll buttons only
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                    ' Scrollbars position
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                    ' Change scrollbar colors
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                End With
            ElseIf numberOfDays = 7 Then
                With Chart1
                    .Series.Add(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value)
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(1, MeanValue(row.Index, 6, 7, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(2, MeanValue(row.Index, 8, 9, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(3, MeanValue(row.Index, 10, 11, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(4, MeanValue(row.Index, 12, 13, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(5, MeanValue(row.Index, 14, 15, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(6, MeanValue(row.Index, 16, 17, EtheCapacity.DataGridView2))
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).Points.AddXY(7, MeanValue(row.Index, 18, 19, EtheCapacity.DataGridView2))

                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).ChartType = SeriesChartType.Line
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value).IsValueShownAsLabel = True
                    .ChartAreas("ChartArea1").AxisX.IsMarginVisible = True
                    .ChartAreas("ChartArea1").AxisY.IsMarginVisible = True
                    .ChartAreas("ChartArea1").AxisY.MajorGrid.Enabled = False
                    .ChartAreas("ChartArea1").AxisX.MajorGrid.Enabled = False
                    .ChartAreas("ChartArea1").Area3DStyle.Enable3D = False
                    .ChartAreas("ChartArea1").Area3DStyle.IsClustered = True
                    .ChartAreas("ChartArea1").AxisX.Title = "Days"
                    .ChartAreas("ChartArea1").AxisY.Title = "Ethernet capacity mean (Mbit/s)"
                    ' Set scrollbar size
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                    ' Show small scroll buttons only
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                    ' Scrollbars position
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                    ' Change scrollbar colors
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.BackColor = Color.LightGray
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonColor = Color.Gray
                    .ChartAreas("ChartArea1").AxisX.ScrollBar.LineColor = Color.Black
                    .Series(EtheCapacity.DataGridView2.Rows(row.Index).Cells(0).Value & "_" & EtheCapacity.DataGridView2.Rows(row.Index).Cells(1).Value)("DrawingStyle") = "Emboss"
                End With
            End If
            'Catch ex As Exception
            'MessageBox.Show(ex.Message & vbLf & "The data contains only non numerical values." & vbLf & "The mean could not be calculated." & vbLf & "Division by zero is impossible!")
            'End Try
        Next
    End Sub
End Class