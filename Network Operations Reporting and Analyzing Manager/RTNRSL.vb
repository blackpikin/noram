
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports Microsoft.Office.Interop


Public Class RTNRSL
    Dim sqlcon As New SqlConnection, sqlcom As New SqlCommand, cstring As String = My.Resources.cstring
    Dim LD As New LoadData
    Dim List As New ListBox, RichText As New RichTextBox, combolist As New ComboBox
    Dim increment As Integer = 0
    Dim combo As New ComboBox
    Dim txtRead As New TextReader
    Dim p As New Printer

    Private Sub RTNRSL_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Home.Show()
    End Sub

    Private Sub RTN_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'set the path of the folder where the files are found
        Try
            List.Items.AddRange(Directory.GetFiles("C:\RTNLinkDump"))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Timer1.Enabled = True
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Me.Close()
        Home.Show()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label9.Text = Date.Now.ToString("MMM dd yyyy   hh:mm:ss")
    End Sub

    Private Sub DrawToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Hide()
        GraphRSLRTN.ShowDialog()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer3.Enabled = False
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
                            Dim row() As String = New String() {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(43), arr(44), arr(45), arr(46)}
                            'add the row to the grid...
                            DataGridView1.Rows.Add(row)
                        End If
                    Next
                    ProgressBar1.PerformStep()
                    'if you already found a date that matches no need to continue looking so exit
                    'count for quick statistics
                    TextBox5.Text = DataGridView1.Rows.Count
                    For Each row As DataGridViewRow In DataGridView1.Rows
                        'conditional formatting
                        If IsNumeric(row.Cells(6).Value) Then
                            If row.Cells(6).Value >= -90 And row.Cells(6).Value <= -56.99 Then
                                row.Cells(6).Style.BackColor = Color.Red
                            ElseIf row.Cells(6).Value >= -56.98 And row.Cells(6).Value <= -51.99 Then
                                row.Cells(6).Style.BackColor = Color.Orange
                            ElseIf row.Cells(6).Value >= -51.98 And row.Cells(6).Value <= -46.99 Then
                                row.Cells(6).Style.BackColor = Color.Yellow
                            ElseIf row.Cells(6).Value >= -46.98 And row.Cells(6).Value <= 0 Then
                                row.Cells(6).Style.BackColor = Color.Green
                            End If
                        End If

                        If IsNumeric(row.Cells(7).Value) Then
                            If row.Cells(7).Value >= -90 And row.Cells(7).Value <= -56.99 Then
                                row.Cells(7).Style.BackColor = Color.Red
                            ElseIf row.Cells(7).Value >= -56.98 And row.Cells(7).Value <= -51.99 Then
                                row.Cells(7).Style.BackColor = Color.Orange
                            ElseIf row.Cells(7).Value >= -51.98 And row.Cells(7).Value <= -46.99 Then
                                row.Cells(7).Style.BackColor = Color.Yellow
                            ElseIf row.Cells(7).Value >= -46.98 And row.Cells(7).Value <= 0 Then
                                row.Cells(7).Style.BackColor = Color.Green
                            End If
                        End If

                        If IsNumeric(row.Cells(8).Value) Then
                            If row.Cells(8).Value >= -90 And row.Cells(8).Value <= -56.99 Then
                                row.Cells(8).Style.BackColor = Color.Red
                            ElseIf row.Cells(8).Value >= -56.98 And row.Cells(8).Value <= -51.99 Then
                                row.Cells(8).Style.BackColor = Color.Orange
                            ElseIf row.Cells(8).Value >= -51.98 And row.Cells(8).Value <= -46.99 Then
                                row.Cells(8).Style.BackColor = Color.Yellow
                            ElseIf row.Cells(8).Value >= -46.98 And row.Cells(8).Value <= 0 Then
                                row.Cells(8).Style.BackColor = Color.Green
                            End If
                        End If

                        If IsNumeric(row.Cells(9).Value) Then
                            If row.Cells(9).Value >= -90 And row.Cells(9).Value <= -56.99 Then
                                row.Cells(9).Style.BackColor = Color.Red
                            ElseIf row.Cells(9).Value >= -56.98 And row.Cells(9).Value <= -51.99 Then
                                row.Cells(9).Style.BackColor = Color.Orange
                            ElseIf row.Cells(9).Value >= -51.98 And row.Cells(9).Value <= -46.99 Then
                                row.Cells(9).Style.BackColor = Color.Yellow
                            ElseIf row.Cells(9).Value >= -46.98 And row.Cells(9).Value <= 0 Then
                                row.Cells(9).Style.BackColor = Color.Green
                            End If
                        End If
                    Next
                    txtRead.createHTMLFile("C:\reportmanager\emailpage.htm", DataGridView1, "RSL Report")
                    WebBrowser1.Navigate("C:\reportmanager\emailpage.htm")
                    Exit Sub
                End If
            Next
            ProgressBar1.PerformStep()
        ElseIf RadioButton2.Checked = True Then
            combo.Items.Clear()
            combolist.Items.Clear()
            'declare strings that hold parts of datetimepicker's date
            Dim dtpMonth, dtpMonth2 As String
            Dim dtpDay, dtpDay2 As String
            'declare a variable to hold the date of the current file
            Dim currentFileDate As String = ""
            'declare the string that receives the selected date from datetimepicker
            Dim selectedDate, selectedDate2 As Date
            'Clear any existing rows on the data grid view
            DataGridView2.Rows.Clear()
            For d As Integer = 0 To 33
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
                MessageBox.Show("The first date must be less than the second date", "Morning Drill", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            If selectedDate = selectedDate2 Then
                MessageBox.Show("The first date must be less than the second date", "Morning Drill", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
                        row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(43), arr(44), arr(45), arr(46)}
                        DataGridView2.Rows.Add(row)
                    End If
                Next
                InsertColumns(1, 10)
                For c As Integer = 14 To 33
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
                        row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(43), arr(44), arr(45), arr(46)}
                        DataGridView2.Rows.Add(row)
                    End If
                Next
                InsertColumns(1, 10)
                InsertColumns(2, 14)
                For c As Integer = 18 To 33
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
                        row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(43), arr(44), arr(45), arr(46)}
                        DataGridView2.Rows.Add(row)
                    End If
                Next
                InsertColumns(1, 10)
                InsertColumns(2, 14)
                InsertColumns(3, 18)
                For c As Integer = 22 To 33
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
                        row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(43), arr(44), arr(45), arr(46)}
                        DataGridView2.Rows.Add(row)
                    End If
                Next
                InsertColumns(1, 10)
                InsertColumns(2, 14)
                InsertColumns(3, 18)
                InsertColumns(4, 22)
                For c As Integer = 26 To 33
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
                        row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(43), arr(44), arr(45), arr(46)}
                        DataGridView2.Rows.Add(row)
                    End If
                Next
                InsertColumns(1, 10)
                InsertColumns(2, 14)
                InsertColumns(3, 18)
                InsertColumns(4, 22)
                InsertColumns(5, 26)
                For c As Integer = 30 To 33
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
                        row = {arr(2), arr(4), arr(7), arr(9), arr(11), arr(13), arr(43), arr(44), arr(45), arr(46)}
                        DataGridView2.Rows.Add(row)
                    End If
                Next
                InsertColumns(1, 10)
                InsertColumns(2, 14)
                InsertColumns(3, 18)
                InsertColumns(4, 22)
                InsertColumns(5, 26)
                InsertColumns(6, 30)
            End If

            DataGridView2.Columns(34).Visible = False
            'calculate averages
            If combo.Items.Count = 2 Then
                For Each rw As DataGridViewRow In DataGridView2.Rows
                    DataGridView2.Rows(rw.Index).Cells(34).Value = Decimal.Round((Val(MeanValue(rw.Index, 8, 9, DataGridView2)) + Val(MeanValue(rw.Index, 12, 13, DataGridView2))) / 2, 2, MidpointRounding.AwayFromZero)
                Next
            ElseIf combo.Items.Count = 3 Then
                For Each rw As DataGridViewRow In DataGridView2.Rows
                    DataGridView2.Rows(rw.Index).Cells(34).Value = Decimal.Round((Val(MeanValue(rw.Index, 8, 9, DataGridView2)) + Val(MeanValue(rw.Index, 12, 13, DataGridView2)) + Val(MeanValue(rw.Index, 16, 17, DataGridView2))) / 3, 2, MidpointRounding.AwayFromZero)
                Next
            ElseIf combo.Items.Count = 4 Then
                For Each rw As DataGridViewRow In DataGridView2.Rows
                    DataGridView2.Rows(rw.Index).Cells(34).Value = Decimal.Round((Val(MeanValue(rw.Index, 8, 9, DataGridView2)) + Val(MeanValue(rw.Index, 12, 13, DataGridView2)) + Val(MeanValue(rw.Index, 16, 17, DataGridView2) + Val(MeanValue(rw.Index, 20, 21, DataGridView2)))) / 4, 2, MidpointRounding.AwayFromZero)
                Next
            ElseIf combo.Items.Count = 5 Then
                For Each rw As DataGridViewRow In DataGridView2.Rows
                    DataGridView2.Rows(rw.Index).Cells(34).Value = Decimal.Round((Val(MeanValue(rw.Index, 8, 9, DataGridView2)) + Val(MeanValue(rw.Index, 12, 13, DataGridView2)) + Val(MeanValue(rw.Index, 16, 17, DataGridView2) + Val(MeanValue(rw.Index, 20, 21, DataGridView2)) + Val(MeanValue(rw.Index, 24, 25, DataGridView2)))) / 5, 2, MidpointRounding.AwayFromZero)
                Next
            ElseIf combo.Items.Count = 6 Then
                For Each rw As DataGridViewRow In DataGridView2.Rows
                    DataGridView2.Rows(rw.Index).Cells(34).Value = Decimal.Round((Val(MeanValue(rw.Index, 8, 9, DataGridView2)) + Val(MeanValue(rw.Index, 12, 13, DataGridView2)) + Val(MeanValue(rw.Index, 16, 17, DataGridView2) + Val(MeanValue(rw.Index, 20, 21, DataGridView2)) + Val(MeanValue(rw.Index, 24, 25, DataGridView2)) + Val(MeanValue(rw.Index, 28, 29, DataGridView2)))) / 6, 2, MidpointRounding.AwayFromZero)
                Next
            ElseIf combo.Items.Count = 7 Then
                For Each rw As DataGridViewRow In DataGridView2.Rows
                    DataGridView2.Rows(rw.Index).Cells(34).Value = Decimal.Round((Val(MeanValue(rw.Index, 8, 9, DataGridView2)) + Val(MeanValue(rw.Index, 12, 13, DataGridView2)) + Val(MeanValue(rw.Index, 16, 17, DataGridView2) + Val(MeanValue(rw.Index, 20, 21, DataGridView2)) + Val(MeanValue(rw.Index, 24, 25, DataGridView2)) + Val(MeanValue(rw.Index, 28, 29, DataGridView2)) + Val(MeanValue(rw.Index, 32, 33, DataGridView2)))) / 7, 2, MidpointRounding.AwayFromZero)
                Next
            End If
            DataGridView2.Columns(34).Visible = True
            '
            'Format colorable cells
            FormatColoredCells(6, 34)

            TextBox5.Text = DataGridView2.Rows.Count
            GroupBox6.Text = "RSL Query for " & combo.Items.Count & " days"
        Else
            MessageBox.Show("Please select a date from view mode!!", "Advice", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
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
                DataGridView2.Rows(x).Cells(c1).Value = lng(43)
                DataGridView2.Rows(x).Cells(c1 + 1).Value = lng(44)
                DataGridView2.Rows(x).Cells(c1 + 2).Value = lng(45)
                DataGridView2.Rows(x).Cells(c1 + 3).Value = lng(46)
            Catch ex As Exception
                row = {lng(2), lng(4), lng(7), lng(9), lng(11), lng(13), lng(43), lng(44), lng(45), lng(46)}
                DataGridView2.Rows.Add(row)

                DataGridView2.Rows(x).Cells(c1).Value = lng(43)
                DataGridView2.Rows(x).Cells(c1 + 1).Value = lng(44)
                DataGridView2.Rows(x).Cells(c1 + 2).Value = lng(45)
                DataGridView2.Rows(x).Cells(c1 + 3).Value = lng(46)
            End Try

        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Timer3.Enabled = False
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.ColumnCount = 10
        DataGridView1.Columns(0).HeaderText = "Source NE Name"
        DataGridView1.Columns(0).Width = 200
        DataGridView1.Columns(1).HeaderText = "Source Board"
        DataGridView1.Columns(2).HeaderText = "Sink NE Name"
        DataGridView1.Columns(3).HeaderText = "Sink Board"
        DataGridView1.Columns(4).HeaderText = "Level"
        DataGridView1.Columns(5).HeaderText = "Source Protect type"
        DataGridView1.Columns(6).HeaderText = "Source NE Max receive power"
        DataGridView1.Columns(7).HeaderText = "Sink NE Max receive power"
        DataGridView1.Columns(8).HeaderText = "Source NE Min receive power"
        DataGridView1.Columns(9).HeaderText = "Sink NE Min receive power"
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' DataGridView1.Rows.Clear()
        For Each row As DataGridViewRow In DataGridView1.Rows
            LD.SaveData("RTNRSL", "SourceNEName, SourceBoard, SinkNEName, SinkBoard, Level, SourceProtectType, SourceMaxRecievePower, SinkMaxRecievePower, SourceMinRecievePower, SinkMinRecievePower", "'" & DataGridView1.Rows(row.Index).Cells(0).Value & "', '" & DataGridView1.Rows(row.Index).Cells(1).Value & "', '" & DataGridView1.Rows(row.Index).Cells(2).Value & "', '" & DataGridView1.Rows(row.Index).Cells(3).Value & "', '" & DataGridView1.Rows(row.Index).Cells(4).Value & "', '" & DataGridView1.Rows(row.Index).Cells(5).Value & "','" & DataGridView1.Rows(row.Index).Cells(6).Value & "','" & DataGridView1.Rows(row.Index).Cells(7).Value & "','" & DataGridView1.Rows(row.Index).Cells(8).Value & "','" & DataGridView1.Rows(row.Index).Cells(9).Value & "'")
        Next
        MessageBox.Show("Query result acknowledged!", "Acknowledging query result", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            Label3.Enabled = True
            Label10.Enabled = True
            GroupBox6.Visible = True
            GroupBox3.Visible = False
        Else
            DateTimePicker1.Enabled = False
            Label10.Enabled = False
            DateTimePicker2.Enabled = False
            Label3.Enabled = False
            GroupBox6.Visible = False
            GroupBox3.Visible = True
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            DateTimePicker1.Enabled = True
        Else
            DateTimePicker1.Enabled = False
        End If

        If RadioButton1.Checked Then
            Label10.Enabled = True
        Else
            Label10.Enabled = False
        End If
    End Sub

    Private Sub EmailToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmailToolStripMenuItem.Click
        'txtRead.WriteText(txtRead.convertGridToHtml(DataGridView1, "RSL Report"), "C:\reportmanager\emailpage.htm")
        ProgressBar1.PerformStep()
        ProgressBar1.PerformStep()
        emailPrompt.mode = "Single RSL"
        emailPrompt.ShowDialog()
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If ProgressBar1.Value = 100 Then
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0
        End If
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        '  Static Lbound As Integer, ubound As Integer
        '  Dim numberofrows = CInt(TextBox5.Text)
        '  Dim remPages As Integer = numberofrows Mod 25
        '  Dim numberofpages As Integer = CInt(numberofrows / 25) + 1
        '  Dim currentpage As Integer = 1
        '  p.Orientation = vbPRORLandscape
        '
        '  If numberofrows < 25 Then
        '      Lbound = 0
        '      ubound = DataGridView1.Rows.Count
        '      PrintPages(Lbound, ubound, currentpage, numberofpages)
        '  ElseIf numberofrows = 25 Then
        '      Lbound = 0
        '      ubound = 25
        '      PrintPages(Lbound, ubound, currentpage, numberofpages)
        '  ElseIf numberofrows > 25 Then
        '      For x As Integer = currentpage To numberofpages
        '          If currentpage = 1 Then
        '              Lbound = 0
        '              ubound = 25
        '              PrintPages(Lbound, ubound, currentpage, numberofpages)
        '          Else
        '              PrintPages(Lbound, ubound, currentpage, numberofpages)
        '          End If
        '          Lbound = Lbound + 25
        '          ubound = ubound + 25
        '          p.NewPage()
        '      Next
        '  End If
        '
        '  p.EndDoc()
    End Sub

    Public Sub PrintPages(ByVal lbound As Integer, ByVal ubound As Integer, ByVal currpage As Integer, ByVal numberofpages As Integer)
        p.Font = New Font("Calibri", 10)
        p.Scale(0, 0, 8.5, 11)

        'draw box
        p.CurrentX = 0.2
        p.CurrentY = 0.5
        p.Line(0.3, 1, 8.2, 10.5, 1, True, False)
        p.CurrentX = 1.7
        p.CurrentY = 1
        p.Line(1.7, 10.5)
        p.CurrentX = 2.5
        p.CurrentY = 1
        p.Line(2.5, 10.5)
        p.CurrentX = 3.9
        p.CurrentY = 1
        p.Line(3.9, 10.5)
        p.CurrentX = 4.5
        p.CurrentY = 1
        p.Line(4.5, 10.5)
        p.CurrentX = 5.5
        p.CurrentY = 1
        p.Line(5.5, 10.5)
        p.CurrentX = 6.3
        p.CurrentY = 1
        p.Line(6.3, 10.5)
        p.CurrentX = 6.8
        p.CurrentY = 1
        p.Line(6.8, 10.5)
        p.CurrentX = 7.3
        p.CurrentY = 1
        p.Line(7.3, 10.5)
        p.CurrentX = 7.8
        p.CurrentY = 1
        p.Line(7.8, 10.5)

        p.CurrentX = 0.3
        p.CurrentY = 1.8
        p.Line(8.2, 1.8)

        p.CurrentX = 0.5
        p.CurrentY = 1
        p.Print("Source NE Name")

        p.CurrentX = 1.8
        p.CurrentY = 1
        p.Print("Source Board")

        p.CurrentX = 2.7
        p.CurrentY = 1
        p.Print("Sink NE Name")

        p.CurrentX = 4
        p.CurrentY = 1
        p.Print("Sink Board")
        p.CurrentX = 4.7
        p.CurrentY = 1
        p.Print("Level")
        p.CurrentX = 5.6
        p.CurrentY = 1
        p.Print("Source Protect")
        p.CurrentX = 5.6
        p.CurrentY = 1.2
        p.Print("type")

        p.CurrentX = 6.4
        p.CurrentY = 1
        p.Print("Source")
        p.CurrentX = 6.4
        p.CurrentY = 1.2
        p.Print("Max")
        p.CurrentX = 6.4
        p.CurrentY = 1.4
        p.Print("Receive")
        p.CurrentX = 6.4
        p.CurrentY = 1.6
        p.Print("Power")

        p.CurrentX = 6.9
        p.CurrentY = 1
        p.Print("Sink")
        p.CurrentX = 6.9
        p.CurrentY = 1.2
        p.Print("Max")
        p.CurrentX = 6.9
        p.CurrentY = 1.4
        p.Print("Receive")
        p.CurrentX = 6.9
        p.CurrentY = 1.6
        p.Print("Power")

        p.CurrentX = 7.4
        p.CurrentY = 1
        p.Print("Source")
        p.CurrentX = 7.4
        p.CurrentY = 1.2
        p.Print("Min")
        p.CurrentX = 7.4
        p.CurrentY = 1.4
        p.Print("Receive")
        p.CurrentX = 7.4
        p.CurrentY = 1.6
        p.Print("Power")

        p.CurrentX = 7.8
        p.CurrentY = 1
        p.Print("Sink")
        p.CurrentX = 7.8
        p.CurrentY = 1.2
        p.Print("Min")
        p.CurrentX = 7.8
        p.CurrentY = 1.4
        p.Print("Receive")
        p.CurrentX = 7.8
        p.CurrentY = 1.6
        p.Print("Power")

        p.CurrentX = 0.5
        p.CurrentY = 1.9
        For i As Integer = lbound To ubound
            p.CurrentX = 0.5
            p.CurrentY = p.CurrentY + 0.1
            p.Print(DataGridView1.Rows(i).Cells(0).Value)
        Next

        p.CurrentX = 1.8
        p.CurrentY = 1.9
        For i As Integer = lbound To ubound
            p.CurrentX = 1.8
            p.CurrentY = p.CurrentY + 0.1
            p.Print(DataGridView1.Rows(i).Cells(1).Value)
        Next

        p.CurrentX = 2.6
        p.CurrentY = 1.9
        For i As Integer = lbound To ubound
            p.CurrentX = 2.6
            p.CurrentY = p.CurrentY + 0.1
            p.Print(DataGridView1.Rows(i).Cells(2).Value)
        Next

        p.CurrentX = 4
        p.CurrentY = 1.9
        For i As Integer = lbound To ubound
            p.CurrentX = 4
            p.CurrentY = p.CurrentY + 0.1
            p.Print(DataGridView1.Rows(i).Cells(3).Value)
        Next

        p.CurrentX = 4.6
        p.CurrentY = 1.9
        For i As Integer = lbound To ubound
            p.CurrentX = 4.6
            p.CurrentY = p.CurrentY + 0.1
            p.Print(DataGridView1.Rows(i).Cells(4).Value)
        Next

        p.CurrentX = 5.6
        p.CurrentY = 1.9
        For i As Integer = lbound To ubound
            p.CurrentX = 5.6
            p.CurrentY = p.CurrentY + 0.1
            p.Print(DataGridView1.Rows(i).Cells(5).Value)
        Next


        p.CurrentX = 6.4
        p.CurrentY = 1.9
        For i As Integer = lbound To ubound
            p.CurrentX = 6.4
            p.CurrentY = p.CurrentY + 0.1
            p.Print(DataGridView1.Rows(i).Cells(6).Value)
        Next

        p.CurrentX = 6.9
        p.CurrentY = 1.9
        For i As Integer = lbound To ubound
            p.CurrentX = 6.9
            p.CurrentY = p.CurrentY + 0.1
            p.Print(DataGridView1.Rows(i).Cells(7).Value)
        Next

        p.CurrentX = 7.4
        p.CurrentY = 1.9
        For i As Integer = lbound To ubound
            p.CurrentX = 7.4
            p.CurrentY = p.CurrentY + 0.1
            p.Print(DataGridView1.Rows(i).Cells(8).Value)
        Next

        p.CurrentX = 7.8
        p.CurrentY = 1.9
        For i As Integer = lbound To ubound
            p.CurrentX = 7.8
            p.CurrentY = p.CurrentY + 0.1
            p.Print(DataGridView1.Rows(i).Cells(9).Value)
        Next

        p.CurrentX = 5
        p.CurrentY = 10.7
        p.Print(currpage & "/" & numberofpages)

    End Sub

    Private Sub PrintOptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintOptionsToolStripMenuItem.Click
        If DataGridView1.Rows.Count <= 0 Then
            MessageBox.Show("Do a query first")
            Exit Sub
        End If
        WebBrowser1.ShowPageSetupDialog()
    End Sub

    Private Sub PrintToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem1.Click
        If DataGridView1.Rows.Count <= 0 Then
            MessageBox.Show("Do a query first")
            Exit Sub
        End If
        WebBrowser1.ShowPrintPreviewDialog()
    End Sub

    Private Sub PrintToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem2.Click
        If DataGridView1.Rows.Count <= 0 Then
            MessageBox.Show("Do a query first")
            Exit Sub
        End If
        'use the webrowser control's print object to print
        WebBrowser1.Print()
    End Sub

    Private Sub FormatColoredCells(ByVal start As Integer, ByVal endvalue As Integer)
        'conditional color formatting
        For Each brow As DataGridViewRow In DataGridView2.Rows
            For i As Integer = start To endvalue
                If IsNumeric(brow.Cells(i).Value) Then
                    If brow.Cells(i).Value >= -90 And brow.Cells(i).Value <= -56.99 Then
                        brow.Cells(i).Style.BackColor = Color.Red
                    ElseIf brow.Cells(i).Value >= -56.98 And brow.Cells(i).Value <= -51.99 Then
                        brow.Cells(i).Style.BackColor = Color.Orange
                    ElseIf brow.Cells(i).Value >= -51.98 And brow.Cells(i).Value <= -46.99 Then
                        brow.Cells(i).Style.BackColor = Color.Yellow
                    ElseIf brow.Cells(i).Value >= -46.98 And brow.Cells(i).Value <= 0 Then
                        brow.Cells(i).Style.BackColor = Color.Green
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub ToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem6.Click
        GraphRSLRTN.Show()
        GraphRSLRTN.Label1.Text = "Bad RSL"
        If combo.Items.Count = 2 Then
            'Well, don't do anything....
        ElseIf combo.Items.Count = 3 Then
            GraphRSLRTN.PlotDI(3)
        ElseIf combo.Items.Count = 4 Then
            GraphRSLRTN.PlotDI(4)
        ElseIf combo.Items.Count = 5 Then
            GraphRSLRTN.PlotDI(5)
        ElseIf combo.Items.Count = 6 Then
            GraphRSLRTN.PlotDI(6)
        ElseIf combo.Items.Count = 7 Then
            GraphRSLRTN.PlotDI(7)
        End If
    End Sub

    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        txtRead.WriteText(txtRead.ConvertRangeGridToHtml(DataGridView2, GroupBox6.Text), "C:\reportmanager\emailpage.htm")
        WebBrowser1.Navigate("C:\reportmanager\emailpage.htm")
        emailPrompt.mode = "Range RSL"
        emailPrompt.ShowDialog()
    End Sub

    ' print sub menu
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
        txtRead.WriteText(txtRead.ConvertRangeGridToHtml(DataGridView2, GroupBox6.Text), "C:\reportmanager\emailpage.htm")
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
        txtRead.WriteText(txtRead.ConvertRangeGridToHtml(DataGridView2, GroupBox6.Text), "C:\reportmanager\emailpage.htm")
        WebBrowser1.Navigate("C:\reportmanager\emailpage.htm")
        ProgressBar1.PerformStep()
        ProgressBar1.PerformStep()
        'use the webrowser control's print object to print
        WebBrowser1.Print()
    End Sub

    '''''''
    Private Function MeanValue(ByVal R As Integer, ByVal C As Integer, ByVal E As Integer, ByVal grid As DataGridView) As Double
        Dim avg As Double = 0.0, n As Integer = 0, sum As Double = 0.0, errMean As String = ""
        Try
            For i As Integer = C To E
                If IsNumeric(grid.Rows(R).Cells(i).Value) Then
                    n = n + 1
                    sum = sum + grid.Rows(R).Cells(i).Value
                End If
            Next
            avg = Decimal.Round(sum / n, 2, MidpointRounding.AwayFromZero)
        Catch ex As Exception
            errMean = "/"
        End Try
        If errMean = "/" Then
            Return Val(errMean)
        Else
            Return avg
        End If
    End Function


    Private Sub CalculateMean(ByVal startVal As Integer, ByVal endVal As Integer)
        For Each row As DataGridViewRow In DataGridView2.Rows
            Try
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
                DataGridView2.Rows(row.Index).Cells(34).Value = mean
            Catch ex As Exception
                DataGridView2.Rows(row.Index).Cells(34).Value = "/"
            End Try
        Next
        DataGridView2.Columns(34).Visible = True
    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        Timer3.Enabled = True
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        'Filter results
        For Each row As DataGridViewRow In DataGridView2.Rows
            Try
                If DataGridView2.Rows(row.Index).Cells(34).Value >= -90 And DataGridView2.Rows(row.Index).Cells(34).Value <= -56.99 Then
                    'Hihihihihi....This is a red value so it stays
                ElseIf DataGridView2.Rows(row.Index).Cells(34).Value >= -56.98 And DataGridView2.Rows(row.Index).Cells(34).Value <= -51.99 Then
                    DataGridView2.Rows.RemoveAt(row.Index)
                ElseIf DataGridView2.Rows(row.Index).Cells(34).Value >= -51.98 And DataGridView2.Rows(row.Index).Cells(34).Value <= -46.99 Then
                    DataGridView2.Rows.RemoveAt(row.Index)
                ElseIf DataGridView2.Rows(row.Index).Cells(34).Value >= -46.98 And DataGridView2.Rows(row.Index).Cells(34).Value <= 0 Then
                    DataGridView2.Rows.RemoveAt(row.Index)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Next
        TextBox5.Text = DataGridView2.Rows.Count
    End Sub

    Private Sub FilteredBadRSLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FilteredBadRSLToolStripMenuItem.Click
        GraphRSLRTN.Show()
        GraphRSLRTN.Label1.Text = "Filtered Bad RSL"
        If combo.Items.Count = 2 Then
            'Well, don't do anything....
        ElseIf combo.Items.Count = 3 Then
            GraphRSLRTN.PlotBadRSL(3)
        ElseIf combo.Items.Count = 4 Then
            GraphRSLRTN.PlotBadRSL(4)
        ElseIf combo.Items.Count = 5 Then
            GraphRSLRTN.PlotBadRSL(5)
        ElseIf combo.Items.Count = 6 Then
            GraphRSLRTN.PlotBadRSL(6)
        ElseIf combo.Items.Count = 7 Then
            GraphRSLRTN.PlotBadRSL(7)
        End If
    End Sub

    Private Sub DeeperInvestigationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeeperInvestigationToolStripMenuItem.Click
       
    End Sub
End Class
