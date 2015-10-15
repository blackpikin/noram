Imports System.IO
Public Class EtheCapacity
    Dim List As New ListBox, RichText As New RichTextBox, combolist As New ComboBox, combo As New ComboBox, txtRead As New TextReader
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Home.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Timer1.Enabled = False
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Label4.Enabled = RadioButton2.Checked
        DateTimePicker2.Enabled = RadioButton2.Checked
        GroupBox4.Visible = RadioButton2.Checked
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.Rows.Clear()
        RichText.Text = ""
        If RadioButton1.Checked = True Then
            ProgressBar1.Visible = True
            ProgressBar1.PerformStep()
            'declare strings that hold parts of datetimepicker's date
            Dim dtpMonth As String
            Dim dtpDay As String
            'declare the string that receives the selected date from datetimepicker
            Dim selectedDate As String
            'Clear any existing rows on the data grid view
            DataGridView1.Rows.Clear()
            'check whether the month is less than 10 and add zero in front
            If DateTimePicker1.Value.Month < 10 Then
                dtpMonth = "0" & DateTimePicker1.Value.Month
            Else
                dtpMonth = DateTimePicker1.Value.Month
            End If
            'check whether the day is less than 10 and add zero infront
            If DateTimePicker1.Value.Day < 10 Then
                dtpDay = "0" & DateTimePicker1.Value.Day
            Else
                dtpDay = DateTimePicker1.Value.Day
            End If

            'set the date string into selectedDate variable
            selectedDate = dtpMonth & "-" & dtpDay & "-" & DateTimePicker1.Value.Year
            ProgressBar1.PerformStep()
            'Iterate thru the items in the listbox which contains the file names
            For i As Integer = 0 To List.Items.Count - 1
                'create an array that contains the pieces of the cut file name using underscore as delimiter
                Dim arrayOfStrings = Split(List.Items.Item(i), "_")
                'test to see if the selected date matches the second item of the cut up file name
                If arrayOfStrings(1) = selectedDate Then
                    'since it matches, create a string that will hold the contents of the file read
                    Dim inputString As String
                    'read the file and store its contents in the inputstring variable
                    Using streamReader As StreamReader = File.OpenText(List.Items.Item(i))
                        inputString = streamReader.ReadToEnd()
                        'transfer the inputstring to richtextbox
                        RichText.Text &= inputString
                    End Using
                    ProgressBar1.PerformStep()
                    'iterate the lines of the richtextbox to
                    For Each line In RichText.Lines
                        'split each line using comma as delimiter
                        Dim arr = Split(line, ",")
                        'get only lines longer than 65 to eliminate empty lines and header lines
                        If arr.Count() > 65 Then
                            'if you want to dynamically create the header's of the datagridview you need to use the line whose count is equal to 65
                            'and supply it to Datagridview.Columns(0).HeaderText = arr(0) and so on
                            'The arr array contains 69 items (i.e 0 - 68) but showing only 0 - 9 for testing
                            'use the values in the array as required
                            Dim avg = 0.0
                            If IsNumeric(arr(53)) And IsNumeric(arr(54)) Then
                                avg = Decimal.Round((CDbl(arr(53)) + CDbl(arr(54))) / 2, 2, MidpointRounding.AwayFromZero)
                            ElseIf IsNumeric(arr(53)) And Not IsNumeric(arr(54)) Then
                                avg = CDbl(arr(53))
                            ElseIf Not IsNumeric(arr(53)) And IsNumeric(arr(54)) Then
                                avg = CDbl(arr(54))
                            Else

                            End If

                            Dim status As String = ""
                            If avg >= 1800 And avg <= 2048 Then
                                status = "Link capacity needs upgrade"
                            ElseIf avg >= 800 And avg <= 1799 Then
                                status = "Monitoring link"
                            ElseIf avg >= 0 And avg <= 799 Then
                                status = "OK"
                            End If

                            Dim row() As String = New String() {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(53), arr(54), avg, status}
                            'add the row to the grid...
                            DataGridView1.Rows.Add(row)
                        End If
                    Next
                    ProgressBar1.PerformStep()
                    'if you already found a date that matches no need to continue looking so exit
                    'count for quick statistics
                    TextBox1.Text = DataGridView1.Rows.Count
                    For Each row As DataGridViewRow In DataGridView1.Rows
                        'conditional formatting
                        If IsNumeric(row.Cells(6).Value) Then
                            If row.Cells(6).Value >= 1800 Then
                                row.Cells(6).Style.BackColor = Color.Red
                            ElseIf row.Cells(6).Value >= 800 And row.Cells(6).Value <= 1799 Then
                                row.Cells(6).Style.BackColor = Color.Orange
                            ElseIf row.Cells(6).Value >= 0 And row.Cells(6).Value <= 799 Then
                                row.Cells(6).Style.BackColor = Color.Green
                            End If
                        End If

                        If IsNumeric(row.Cells(7).Value) Then
                            If row.Cells(7).Value >= 1800 Then
                                row.Cells(7).Style.BackColor = Color.Red
                            ElseIf row.Cells(7).Value >= 800 And row.Cells(7).Value <= 1799 Then
                                row.Cells(7).Style.BackColor = Color.Orange
                            ElseIf row.Cells(7).Value >= 0 And row.Cells(7).Value <= 799 Then
                                row.Cells(7).Style.BackColor = Color.Green
                            End If
                        End If

                        If IsNumeric(row.Cells(8).Value) Then
                            If row.Cells(8).Value >= 1800 Then
                                row.Cells(8).Style.BackColor = Color.Red
                            ElseIf row.Cells(8).Value >= 800 And row.Cells(8).Value <= 1799 Then
                                row.Cells(8).Style.BackColor = Color.Orange
                            ElseIf row.Cells(8).Value >= 0 And row.Cells(8).Value <= 799 Then
                                row.Cells(8).Style.BackColor = Color.Green
                            End If
                        End If

                        If IsNumeric(row.Cells(9).Value) Then
                            If row.Cells(9).Value >= 1800 Then
                                row.Cells(9).Style.BackColor = Color.Red
                            ElseIf row.Cells(9).Value >= 800 And row.Cells(9).Value <= 1799 Then
                                row.Cells(9).Style.BackColor = Color.Orange
                            ElseIf row.Cells(9).Value >= 0 And row.Cells(9).Value <= 799 Then
                                row.Cells(9).Style.BackColor = Color.Green
                            End If
                        End If
                    Next
                    txtRead.createHTMLFile("C:\reportmanager\emailpage.htm", DataGridView1, "Ethernet Capacity Report")
                    ProgressBar1.PerformStep()
                    ProgressBar1.Visible = False
                    WebBrowser1.Navigate("C:\reportmanager\emailpage.htm")
                    Exit Sub
                End If
            Next
        End If

        If RadioButton2.Checked = True Then
            Timer1.Enabled = False
            QueryRange()
        End If
    End Sub

    Private Sub EtheCapacity_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Icon = Home.Icon
        'set the path of the folder where the files are found
        Try
            List.Items.AddRange(Directory.GetFiles("C:\RTNLinkDump"))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub PrintOptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintOptionsToolStripMenuItem.Click
        If DataGridView1.Rows.Count <= 0 Then
            MessageBox.Show("Do a query first")
            Exit Sub
        End If
        WebBrowser1.ShowPageSetupDialog()
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        If DataGridView1.Rows.Count <= 0 Then
            MessageBox.Show("Do a query first")
            Exit Sub
        End If
        WebBrowser1.ShowPrintPreviewDialog()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        If DataGridView1.Rows.Count <= 0 Then
            MessageBox.Show("Do a query first")
            Exit Sub
        End If
        'use the webrowser control's print object to print
        WebBrowser1.Print()
    End Sub

    Private Sub EmailQueryResultToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmailQueryResultToolStripMenuItem.Click
        ProgressBar1.PerformStep()
        ProgressBar1.PerformStep()
        emailPrompt.mode = "Single Ethernet"
        emailPrompt.ShowDialog()
    End Sub

    Private Sub QueryRange()
        combo.Items.Clear()
        combolist.Items.Clear()
        DataGridView2.Rows.Clear()
        RichText.Text = ""
        'declare strings that hold parts of datetimepicker's date
        Dim dtpMonth, dtpMonth2 As String
        Dim dtpDay, dtpDay2 As String
        'declare a variable to hold the date of the current file
        Dim currentFileDate As String = ""
        'declare the string that receives the selected date from datetimepicker
        Dim selectedDate, selectedDate2 As Date
        'Clear any existing rows on the data grid view
        DataGridView2.Rows.Clear()
        For d As Integer = 0 To 21
            DataGridView2.Columns(d).Visible = True
        Next
        'check whether the month is less than 10 and add zero in front
        If DateTimePicker1.Value.Month < 10 Then
            dtpMonth = "0" & DateTimePicker1.Value.Month
        Else
            dtpMonth = DateTimePicker1.Value.Month
        End If
        'check whether the day is less than 10 and add zero infront
        If DateTimePicker1.Value.Day < 10 Then
            dtpDay = "0" & DateTimePicker1.Value.Day
        Else
            dtpDay = DateTimePicker1.Value.Day
        End If

        If DateTimePicker2.Value.Month < 10 Then
            dtpMonth2 = "0" & DateTimePicker2.Value.Month
        Else
            dtpMonth2 = DateTimePicker2.Value.Month
        End If
        'check whether the day is less than 10 and add zero infront
        If DateTimePicker2.Value.Day < 10 Then
            dtpDay2 = "0" & DateTimePicker2.Value.Day
        Else
            dtpDay2 = DateTimePicker2.Value.Day
        End If

        'set the date string into selectedDate variable
        selectedDate = dtpMonth & "-" & dtpDay & "-" & DateTimePicker1.Value.Year
        'set the date string into selectedDate2 variable
        selectedDate2 = dtpMonth2 & "-" & dtpDay2 & "-" & DateTimePicker2.Value.Year

        'test to make sure the selected dates are a range
        If selectedDate > selectedDate2 Then
            MessageBox.Show("The first date must be less than the second date", "NORAM", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If selectedDate = selectedDate2 Then
            MessageBox.Show("The first date must be less than the second date", "NORAM", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        'Iterate thru the items in the listbox which contains the file names
        For i As Integer = 0 To List.Items.Count - 1
            'create an array that contains the pieces of the cut file name using underscore as delimiter
            Dim arrayOfStrings = Split(List.Items.Item(i), "_")
            'test to see if the selected date matches the second item of the cut up file name
            If (arrayOfStrings(1) >= selectedDate) And (arrayOfStrings(1) <= selectedDate2) Then
                If Not combo.Items.Contains(arrayOfStrings(1)) Then
                    combo.Items.Add(arrayOfStrings(1))
                    combolist.Items.Add(List.Items.Item(i))
                End If
            End If
        Next
        'ensure that the user did not select a range longer than 7 days
        If combo.Items.Count > 7 Then
            MessageBox.Show("The range is too long." & vbLf & "You can only select a maximum of 7 days.", "Morning Drill", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            combo.Items.Clear()
            combolist.Items.Clear()
            Exit Sub
        End If

        If combo.Items.Count < 3 Then
            MessageBox.Show("The range is too short." & vbLf & "You can only select a maximum of 3 days.", "Morning Drill", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            combo.Items.Clear()
            combolist.Items.Clear()
            Exit Sub
        End If

        Dim inputString As String
        Dim row() As String = New String() {}

        If combo.Items.Count = 2 Then
            RichText.Text = ""
            Using streamReader As StreamReader = File.OpenText(combolist.Items(0))
                inputString = streamReader.ReadToEnd()
                RichText.Text &= inputString
            End Using
            For Each line In RichText.Lines
                Dim arr = Split(line, ",")
                If arr.Count() > 65 Then
                    row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(53), arr(54), "", ""}
                    DataGridView2.Rows.Add(row)
                End If
            Next
            InsertColumns(1, 8)
            For c As Integer = 10 To 21
                DataGridView2.Columns(c).Visible = False
            Next
        ElseIf combo.Items.Count = 3 Then
            RichText.Text = ""
            Using streamReader As StreamReader = File.OpenText(combolist.Items(0))
                inputString = streamReader.ReadToEnd()
                RichText.Text &= inputString
            End Using
            For Each line In RichText.Lines
                Dim arr = Split(line, ",")
                If arr.Count() > 65 Then
                    row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(53), arr(54), "", ""}
                    DataGridView2.Rows.Add(row)
                End If
            Next
            InsertColumns(1, 8)
            InsertColumns(2, 10)
            For c As Integer = 12 To 21
                DataGridView2.Columns(c).Visible = False
            Next
        ElseIf combo.Items.Count = 4 Then
            RichText.Text = ""
            Using streamReader As StreamReader = File.OpenText(combolist.Items(0))
                inputString = streamReader.ReadToEnd()
                RichText.Text &= inputString
            End Using
            For Each line In RichText.Lines
                Dim arr = Split(line, ",")
                If arr.Count() > 65 Then
                    row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(53), arr(54), "", ""}
                    DataGridView2.Rows.Add(row)
                End If
            Next
            InsertColumns(1, 8)
            InsertColumns(2, 10)
            InsertColumns(3, 12)
            For c As Integer = 14 To 21
                DataGridView2.Columns(c).Visible = False
            Next
        ElseIf combo.Items.Count = 5 Then
            RichText.Text = ""
            Using streamReader As StreamReader = File.OpenText(combolist.Items(0))
                inputString = streamReader.ReadToEnd()
                RichText.Text &= inputString
            End Using
            For Each line In RichText.Lines
                Dim arr = Split(line, ",")
                If arr.Count() > 65 Then
                    row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(53), arr(54), "", ""}
                    DataGridView2.Rows.Add(row)
                End If
            Next
            InsertColumns(1, 8)
            InsertColumns(2, 10)
            InsertColumns(3, 12)
            InsertColumns(4, 14)
            For c As Integer = 16 To 21
                DataGridView2.Columns(c).Visible = False
            Next
        ElseIf combo.Items.Count = 6 Then
            RichText.Text = ""
            Using streamReader As StreamReader = File.OpenText(combolist.Items(0))
                inputString = streamReader.ReadToEnd()
                RichText.Text &= inputString
            End Using
            For Each line In RichText.Lines
                Dim arr = Split(line, ",")
                If arr.Count() > 65 Then
                    row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(53), arr(54), "", ""}
                    DataGridView2.Rows.Add(row)
                End If
            Next
            InsertColumns(1, 8)
            InsertColumns(2, 10)
            InsertColumns(3, 12)
            InsertColumns(4, 14)
            InsertColumns(5, 16)
            For c As Integer = 18 To 21
                DataGridView2.Columns(c).Visible = False
            Next
        ElseIf combo.Items.Count = 7 Then
            RichText.Text = ""
            Using streamReader As StreamReader = File.OpenText(combolist.Items(0))
                inputString = streamReader.ReadToEnd()
                RichText.Text &= inputString
            End Using
            For Each line In RichText.Lines
                Dim arr = Split(line, ",")
                If arr.Count() > 65 Then
                    row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(53), arr(54), "", ""}
                    DataGridView2.Rows.Add(row)
                End If
            Next
            InsertColumns(1, 8)
            InsertColumns(2, 10)
            InsertColumns(3, 12)
            InsertColumns(4, 14)
            InsertColumns(5, 16)
            InsertColumns(6, 18)
        End If

        DataGridView2.Columns(20).Visible = False
        DataGridView2.Columns(21).Visible = False
        'calculate averages
        If combo.Items.Count = 2 Then
            CalculateMean(6, 9)
        ElseIf combo.Items.Count = 3 Then
            CalculateMean(6, 11)
        ElseIf combo.Items.Count = 4 Then
            CalculateMean(6, 13)
        ElseIf combo.Items.Count = 5 Then
            CalculateMean(6, 15)
        ElseIf combo.Items.Count = 6 Then
            CalculateMean(6, 17)
        ElseIf combo.Items.Count = 7 Then
            CalculateMean(6, 19)
        End If

        '
        'Format colorable cells
        FormatColoredCells(6, 20)

        TextBox1.Text = DataGridView2.Rows.Count
        GroupBox4.Text = "Ethernet capacity Query for " & combo.Items.Count & " days"
    End Sub

    Private Sub InsertColumns(ByVal combolistItem As Integer, ByVal c1 As Integer)
        RichText.Text = ""
        Dim inputString As String
        Dim row() As String = New String() {}
        Using streamReader As StreamReader = File.OpenText(combolist.Items(combolistItem))
            inputString = streamReader.ReadToEnd()
            RichText.Text &= inputString
        End Using
        Dim RT As New ListBox
        For Each line In RichText.Lines
            Dim arr = Split(line, ",")
            If arr.Count() > 65 Then
                RT.Items.Add(line)
            End If
        Next

        For x As Integer = 0 To RT.Items.Count - 1
            Dim lng = Split(RT.Items.Item(x), ",")
            Try
                DataGridView2.Rows(x).Cells(c1).Value = lng(53)
                DataGridView2.Rows(x).Cells(c1 + 1).Value = lng(54)
            Catch ex As Exception
                row = {lng(2), lng(4), lng(7), lng(9), lng(11), lng(13), lng(53), lng(54), "", ""}
                DataGridView2.Rows.Add(row)

                DataGridView2.Rows(x).Cells(c1).Value = lng(53)
                DataGridView2.Rows(x).Cells(c1 + 1).Value = lng(54)
            End Try

        Next

    End Sub

    Private Sub FormatColoredCells(ByVal start As Integer, ByVal endvalue As Integer)
        'conditional color formatting
        For Each brow As DataGridViewRow In DataGridView2.Rows
            For i As Integer = start To endvalue
                If IsNumeric(brow.Cells(i).Value) Then
                    If brow.Cells(i).Value >= 1800 Then
                        brow.Cells(i).Style.BackColor = Color.Red
                    ElseIf brow.Cells(i).Value >= 800 And brow.Cells(i).Value <= 1799 Then
                        brow.Cells(i).Style.BackColor = Color.Orange
                    ElseIf brow.Cells(i).Value >= 0 And brow.Cells(i).Value <= 799 Then
                        brow.Cells(i).Style.BackColor = Color.Green
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub CalculateMean(ByVal startVal As Integer, ByVal endVal As Integer)
        For Each row As DataGridViewRow In DataGridView2.Rows
            Dim sum As Double = 0.0, colCount As Integer = 0
            For i As Integer = startVal To endVal
                If IsNumeric(DataGridView2.Rows(row.Index).Cells(i).Value) Then
                    sum = DataGridView2.Rows(row.Index).Cells(i).Value + sum
                    colCount = colCount + 1
                Else
                    sum = 0.0 + sum
                End If
            Next
            Dim mean = Decimal.Round(sum / colCount, 2, MidpointRounding.AwayFromZero)
            DataGridView2.Rows(row.Index).Cells(20).Value = mean
            DataGridView2.Rows(row.Index).Cells(21).Value = FormulateStatus(mean)
        Next
        DataGridView2.Columns(20).Visible = True
        DataGridView2.Columns(21).Visible = True
    End Sub

    Private Function FormulateStatus(ByVal mean As Double) As String
        Dim status As String = ""
        If mean >= 1800 Then
            status = "Link capacity needs upgrade"
        ElseIf mean >= 800 And mean <= 1799 Then
            status = "Monitoring..."
        ElseIf mean >= 0 And mean <= 799 Then
            status = "OK"
        End If
        Return status
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Filter results
        For Each row As DataGridViewRow In DataGridView2.Rows
            Try
                If DataGridView2.Rows(row.Index).Cells(20).Value >= 1800 Then
                    'Hihihihihi....This is a red value so it stays
                ElseIf DataGridView2.Rows(row.Index).Cells(20).Value >= 800 And DataGridView2.Rows(row.Index).Cells(20).Value <= 1799 Then
                    DataGridView2.Rows.RemoveAt(row.Index)
                ElseIf DataGridView2.Rows(row.Index).Cells(20).Value >= 0 And DataGridView2.Rows(row.Index).Cells(20).Value <= 799 Then
                    DataGridView2.Rows.RemoveAt(row.Index)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Next
        TextBox1.Text = DataGridView2.Rows.Count
    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        Timer1.Enabled = True
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        If DataGridView2.Rows.Count <= 0 Then
            MessageBox.Show("Do a query first")
            Exit Sub
        End If
        WebBrowser1.ShowPageSetupDialog()
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        If DataGridView2.Rows.Count <= 0 Then
            MessageBox.Show("Do a query first")
            Exit Sub
        End If
        txtRead.WriteText(txtRead.ConvertRangeGridToHtml(DataGridView2, GroupBox4.Text), "C:\reportmanager\emailpage.htm")
        WebBrowser1.Navigate("C:\reportmanager\emailpage.htm")
        ProgressBar1.PerformStep()
        ProgressBar1.PerformStep()
        WebBrowser1.ShowPrintPreviewDialog()
    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        If DataGridView2.Rows.Count <= 0 Then
            MessageBox.Show("Do a query first")
            Exit Sub
        End If
        txtRead.WriteText(txtRead.ConvertRangeGridToHtml(DataGridView2, GroupBox4.Text), "C:\reportmanager\emailpage.htm")
        WebBrowser1.Navigate("C:\reportmanager\emailpage.htm")
        ProgressBar1.PerformStep()
        ProgressBar1.PerformStep()
        WebBrowser1.Print()
    End Sub

    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        ProgressBar1.PerformStep()
        ProgressBar1.PerformStep()
        emailPrompt.mode = "Range Ethernet"
        emailPrompt.ShowDialog()
    End Sub

    Private Sub FilteredBadRSLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FilteredBadRSLToolStripMenuItem.Click
        GraphRSLRTN.Show()
        GraphRSLRTN.Label1.Text = "Filtered Bad Ethernet Capacity"
        If combo.Items.Count = 2 Then
            'Well, don't do anything....
        ElseIf combo.Items.Count = 3 Then
            GraphRSLRTN.PlotBadEthernet(3)
        ElseIf combo.Items.Count = 4 Then
            GraphRSLRTN.PlotBadEthernet(4)
        ElseIf combo.Items.Count = 5 Then
            GraphRSLRTN.PlotBadEthernet(5)
        ElseIf combo.Items.Count = 6 Then
            GraphRSLRTN.PlotBadEthernet(6)
        ElseIf combo.Items.Count = 7 Then
            GraphRSLRTN.PlotBadEthernet(7)
        End If
    End Sub
End Class