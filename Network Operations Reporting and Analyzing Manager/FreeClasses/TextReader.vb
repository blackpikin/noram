Imports System.IO
Imports System.Data.SqlClient

Public Class TextReader
    Dim comp As String = ""
    Dim add1 As String = ""
    Dim add2 As String = ""
    Dim tel1 As String = ""
    Dim tel2 As String = ""
    Dim email As String = ""
    Public Sub WriteText(ByVal ctrl As Control, ByVal path As String)
        Using streamWriter As StreamWriter = File.CreateText(path)
            streamWriter.WriteLine(ctrl.Text)
        End Using
    End Sub

    Public Sub WriteText(ByVal txt As String, ByVal path As String)
        Using streamWriter As StreamWriter = File.CreateText(path)
            streamWriter.WriteLine(txt)
        End Using
    End Sub

    Public Sub AppendText(ByVal ctrl As Control, ByVal path As String)
        Using streamWriter As StreamWriter = File.AppendText(path)
            streamWriter.WriteLine(ctrl.Text)
        End Using
    End Sub

    Public Sub ReadText(ByVal ctrl As Control, ByVal path As String)
        Dim inputString As String
        ctrl.Text = ""
        Using streamReader As StreamReader = File.OpenText(path)
            inputString = streamReader.ReadLine()
            While (inputString <> Nothing)
                ctrl.Text &= inputString & vbLf
                inputString = streamReader.ReadLine()
            End While
        End Using
    End Sub

    Public Sub createHTMLFile(ByVal path As String, ByVal grid As DataGridView, ByVal title As String)
        Try
            Using streamWriter As StreamWriter = File.CreateText(path)
                streamWriter.WriteLine("<!DOCTYPE HTML PUBLIC -//W3C//DTD HTML 4.0 Transitional//EN>")
                streamWriter.WriteLine("<html>")
                streamWriter.WriteLine("<head>")
                streamWriter.WriteLine("<title>" & title & "</title>")
                streamWriter.WriteLine("<style type=""text/css""" & ">" & "")
                'create style 1
                streamWriter.WriteLine(".style1")
                streamWriter.WriteLine("{")
                streamWriter.WriteLine("width: 514px;")
                streamWriter.WriteLine("}")

                'create style 2
                streamWriter.WriteLine(".style2")
                streamWriter.WriteLine("{")
                streamWriter.WriteLine("width: 216px;")
                streamWriter.WriteLine("font-weight: bold;")
                streamWriter.WriteLine("}")
                streamWriter.WriteLine("</style>")

                Dim richtext As New RichTextBox, combo As New ComboBox
                ReadText(richtext, "C:\reportmanager\facilityinfo.txt")
                For Each line In richtext.Lines
                    combo.Items.Add(line)
                Next

                streamWriter.WriteLine("</head>")
                streamWriter.WriteLine("<body>")
                'streamWriter.WriteLine("<img alt="""" src=""E:\VB projects\MorningDrill\telecom_250x250.jpg""" & " style=""width: 290px; height: 177px""/>")
                streamWriter.WriteLine("<h1>" & combo.Items(0) & "</h1>")
                streamWriter.WriteLine("<h3>" & combo.Items(1) & "&nbsp" & combo.Items(2) & "&nbsp" & combo.Items(3) & "</h3>")
                streamWriter.WriteLine("<h3>Tel:" & combo.Items(4) & "&nbsp; Email:" & combo.Items(5) & "</h3>")
                'start table1
                streamWriter.WriteLine("<table style=" & """" & "width: 100%;""" & " border=" & """1""" & ">")

                'first row of the table
                streamWriter.WriteLine("<tr>")

                'first column of the table
                streamWriter.WriteLine("<td>")
                streamWriter.WriteLine("<strong>" & grid.Columns(0).HeaderText & "</strong>")
                streamWriter.WriteLine("</td>")

                'second column of the table
                streamWriter.WriteLine("<td>")
                streamWriter.WriteLine("<strong>" & grid.Columns(1).HeaderText & "</strong>")
                streamWriter.WriteLine("</td>")

                'third column
                streamWriter.WriteLine("<td>")
                streamWriter.WriteLine("<strong>" & grid.Columns(2).HeaderText & "</strong>")
                streamWriter.WriteLine("</td>")

                'fourth column
                streamWriter.WriteLine("<td>")
                streamWriter.WriteLine("<strong>" & grid.Columns(3).HeaderText & "</strong>")
                streamWriter.WriteLine("</td>")

                'fifth column
                streamWriter.WriteLine("<td>")
                streamWriter.WriteLine("<strong>" & grid.Columns(4).HeaderText & "</strong>")
                streamWriter.WriteLine("</td>")

                streamWriter.WriteLine("<td>")
                streamWriter.WriteLine("<strong>" & grid.Columns(5).HeaderText & "</strong>")
                streamWriter.WriteLine("</td>")

                streamWriter.WriteLine("<td>")
                streamWriter.WriteLine("<strong>" & grid.Columns(6).HeaderText & "</strong>")
                streamWriter.WriteLine("</td>")

                streamWriter.WriteLine("<td>")
                streamWriter.WriteLine("<strong>" & grid.Columns(7).HeaderText & "</strong>")
                streamWriter.WriteLine("</td>")

                streamWriter.WriteLine("<td>")
                streamWriter.WriteLine("<strong>" & grid.Columns(8).HeaderText & "</strong>")
                streamWriter.WriteLine("</td>")

                streamWriter.WriteLine("<td>")
                streamWriter.WriteLine("<strong>" & grid.Columns(9).HeaderText & "</strong>")
                streamWriter.WriteLine("</td>")



                'end first row of the table
                streamWriter.WriteLine("</tr>")

                For Each row As DataGridViewRow In grid.Rows
                    streamWriter.WriteLine("<tr>")
                    'first column of the table
                    streamWriter.WriteLine("<td>")
                    streamWriter.WriteLine(grid.Rows(row.Index).Cells(0).Value)
                    streamWriter.WriteLine("</td>")

                    'second column of the table
                    streamWriter.WriteLine("<td>")
                    streamWriter.WriteLine(grid.Rows(row.Index).Cells(1).Value)
                    streamWriter.WriteLine("</td>")

                    'third column
                    streamWriter.WriteLine("<td>")
                    streamWriter.WriteLine(grid.Rows(row.Index).Cells(2).Value)
                    streamWriter.WriteLine("</td>")

                    'fourth column
                    streamWriter.WriteLine("<td>")
                    streamWriter.WriteLine(grid.Rows(row.Index).Cells(3).Value)
                    streamWriter.WriteLine("</td>")

                    'fifth column
                    streamWriter.WriteLine("<td>")
                    streamWriter.WriteLine(grid.Rows(row.Index).Cells(4).Value)
                    streamWriter.WriteLine("</td>")

                    'fifth column
                    streamWriter.WriteLine("<td>")
                    streamWriter.WriteLine(grid.Rows(row.Index).Cells(5).Value)
                    streamWriter.WriteLine("</td>")

                    'sixth column
                    Dim bgcolor As String = "white;"
                    If row.Cells(6).Style.BackColor = Color.Red Then
                        bgcolor = "#ff0000;"
                    ElseIf row.Cells(6).Style.BackColor = Color.Orange Then
                        bgcolor = "orange;"
                    ElseIf row.Cells(6).Style.BackColor = Color.Yellow Then
                        bgcolor = "yellow;"
                    ElseIf row.Cells(6).Style.BackColor = Color.Green Then
                        bgcolor = "green;"
                    End If

                    streamWriter.WriteLine("<td style='" & "background-color:" & bgcolor & "'>")
                    streamWriter.WriteLine(grid.Rows(row.Index).Cells(6).Value)
                    streamWriter.WriteLine("</td>")

                    'seventh column
                    If row.Cells(7).Style.BackColor = Color.Red Then
                        bgcolor = "#ff0000;"
                    ElseIf row.Cells(7).Style.BackColor = Color.Orange Then
                        bgcolor = "orange;"
                    ElseIf row.Cells(7).Style.BackColor = Color.Yellow Then
                        bgcolor = "yellow;"
                    ElseIf row.Cells(7).Style.BackColor = Color.Green Then
                        bgcolor = "green;"
                    End If
                    streamWriter.WriteLine("<td style='" & "background-color:" & bgcolor & "'>")
                    streamWriter.WriteLine(grid.Rows(row.Index).Cells(7).Value)
                    streamWriter.WriteLine("</td>")

                    'eighth column
                    If row.Cells(8).Style.BackColor = Color.Red Then
                        bgcolor = "#ff0000;"
                    ElseIf row.Cells(8).Style.BackColor = Color.Orange Then
                        bgcolor = "orange;"
                    ElseIf row.Cells(8).Style.BackColor = Color.Yellow Then
                        bgcolor = "yellow;"
                    ElseIf row.Cells(8).Style.BackColor = Color.Green Then
                        bgcolor = "green;"
                    End If
                    streamWriter.WriteLine("<td style='" & "background-color:" & bgcolor & "'>")
                    streamWriter.WriteLine(grid.Rows(row.Index).Cells(8).Value)
                    streamWriter.WriteLine("</td>")

                    'ninth column
                    If row.Cells(9).Style.BackColor = Color.Red Then
                        bgcolor = "#ff0000;"
                    ElseIf row.Cells(9).Style.BackColor = Color.Orange Then
                        bgcolor = "orange;"
                    ElseIf row.Cells(9).Style.BackColor = Color.Yellow Then
                        bgcolor = "yellow;"
                    ElseIf row.Cells(9).Style.BackColor = Color.Green Then
                        bgcolor = "green;"
                    End If
                    streamWriter.WriteLine("<td style='" & "background-color:" & bgcolor & "'>")
                    streamWriter.WriteLine(grid.Rows(row.Index).Cells(9).Value)
                    streamWriter.WriteLine("</td>")

                    'end first row of the table
                    streamWriter.WriteLine("</tr>")

                Next
                'end table1
                streamWriter.WriteLine("</table>")
                streamWriter.WriteLine("<h4>" & combo.Items(6) & "</h4>")
                streamWriter.WriteLine("</body>")
                streamWriter.WriteLine("</html>")
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Function convertGridToHtml(ByVal grid As DataGridView, ByVal title As String) As String
        RTNRSL.ProgressBar1.Visible = True
        RTNRSL.ProgressBar1.PerformStep()
        Dim output As String = "" ' a string to be used as the HTML output
        output &= "<!DOCTYPE HTML PUBLIC -//W3C//DTD HTML 4.0 Transitional//EN>"
        output &= "<html>"
        output &= "<head>"
        output &= "<title>" & title & "</title>"
        output &= "<style type=""text/css""" & ">" & ""
        'create style 1
        output &= ".style1"
        output &= "{"
        output &= "width: 514px;"
        output &= "}"

        'create style 2
        output &= ".style2"
        output &= "{"
        output &= "width: 216px;"
        output &= "font-weight: bold;"
        output &= "}"
        output &= "</style>"

        Dim richtext As New RichTextBox, combo As New ComboBox
        ReadText(richtext, "C:\reportmanager\facilityinfo.txt")
        For Each line In richtext.Lines
            combo.Items.Add(line)
        Next

        output &= "</head>"
        output &= "<body>"
        'streamWriter.WriteLine("<img alt="""" src=""E:\VB projects\MorningDrill\telecom_250x250.jpg""" & " style=""width: 290px; height: 177px""/>")
        output &= "<h1>" & combo.Items(0) & "</h1>"
        output &= "<h3>" & combo.Items(1) & "&nbsp" & combo.Items(2) & "&nbsp" & combo.Items(3) & "</h3>"
        output &= "<h3>Tel:" & combo.Items(4) & "&nbsp; Email:" & combo.Items(5) & "</h3>"
        'start table1
        output &= "<table style=" & """" & "width: 100%;""" & " border=" & """1""" & ">"

        'first row of the table
        output &= "<tr>"
        RTNRSL.ProgressBar1.PerformStep()
        'first column of the table
        output &= "<td>"
        output &= "<strong>" & grid.Columns(0).HeaderText & "</strong>"
        output &= "</td>"

        'second column of the table
        output &= "<td>"
        output &= "<strong>" & grid.Columns(1).HeaderText & "</strong>"
        output &= "</td>"

        'third column
        output &= "<td>"
        output &= "<strong>" & grid.Columns(2).HeaderText & "</strong>"
        output &= "</td>"

        output &= "<td>"
        output &= "<strong>" & grid.Columns(3).HeaderText & "</strong>"
        output &= "</td>"

        output &= "<td>"
        output &= "<strong>" & grid.Columns(4).HeaderText & "</strong>"
        output &= "</td>"

        output &= "<td>"
        output &= "<strong>" & grid.Columns(5).HeaderText & "</strong>"
        output &= "</td>"

        output &= "<td>"
        output &= "<strong>" & grid.Columns(6).HeaderText & "</strong>"
        output &= "</td>"

        output &= "<td>"
        output &= "<strong>" & grid.Columns(7).HeaderText & "</strong>"
        output &= "</td>"

        output &= "<td>"
        output &= "<strong>" & grid.Columns(8).HeaderText & "</strong>"
        output &= "</td>"

        output &= "<td>"
        output &= "<strong>" & grid.Columns(9).HeaderText & "</strong>"
        output &= "</td>"

        'end first row of the table
        output &= "</tr>"

        For Each row As DataGridViewRow In grid.Rows
            output &= "<tr>"
            output &= "<td>"
            output &= grid.Rows(row.Index).Cells(0).Value
            output &= "</td>"
            output &= "<td>"
            output &= grid.Rows(row.Index).Cells(1).Value
            output &= "</td>"
            output &= "<td>"
            output &= grid.Rows(row.Index).Cells(2).Value
            output &= "</td>"
            output &= "<td>"
            output &= grid.Rows(row.Index).Cells(3).Value
            output &= "</td>"

            'fifth column
            output &= "<td>"
            output &= grid.Rows(row.Index).Cells(4).Value
            output &= "</td>"

            'fifth column
            output &= "<td>"
            output &= grid.Rows(row.Index).Cells(5).Value
            output &= "</td>"

            'sixth column
            Dim bgcolor As String = "white;"
            If row.Cells(6).Style.BackColor = Color.Red Then
                bgcolor = "#ff0000;"
            ElseIf row.Cells(6).Style.BackColor = Color.Orange Then
                bgcolor = "orange;"
            ElseIf row.Cells(6).Style.BackColor = Color.Yellow Then
                bgcolor = "yellow;"
            ElseIf row.Cells(6).Style.BackColor = Color.Green Then
                bgcolor = "green;"
            End If

            output &= "<td style='" & "background-color:" & bgcolor & "'>"
            output &= grid.Rows(row.Index).Cells(6).Value
            output &= "</td>"

            RTNRSL.ProgressBar1.PerformStep()

            'seventh column
            If row.Cells(7).Style.BackColor = Color.Red Then
                bgcolor = "#ff0000;"
            ElseIf row.Cells(7).Style.BackColor = Color.Orange Then
                bgcolor = "orange;"
            ElseIf row.Cells(7).Style.BackColor = Color.Yellow Then
                bgcolor = "yellow;"
            ElseIf row.Cells(7).Style.BackColor = Color.Green Then
                bgcolor = "green;"
            End If
            output &= "<td style='" & "background-color:" & bgcolor & "'>"
            output &= grid.Rows(row.Index).Cells(7).Value
            output &= "</td>"

            'eighth column
            If row.Cells(8).Style.BackColor = Color.Red Then
                bgcolor = "#ff0000;"
            ElseIf row.Cells(8).Style.BackColor = Color.Orange Then
                bgcolor = "orange;"
            ElseIf row.Cells(8).Style.BackColor = Color.Yellow Then
                bgcolor = "yellow;"
            ElseIf row.Cells(8).Style.BackColor = Color.Green Then
                bgcolor = "green;"
            End If
            output &= "<td style='" & "background-color:" & bgcolor & "'>"
            output &= grid.Rows(row.Index).Cells(8).Value
            output &= "</td>"

            'ninth column
            If row.Cells(9).Style.BackColor = Color.Red Then
                bgcolor = "#ff0000;"
            ElseIf row.Cells(9).Style.BackColor = Color.Orange Then
                bgcolor = "orange;"
            ElseIf row.Cells(9).Style.BackColor = Color.Yellow Then
                bgcolor = "yellow;"
            ElseIf row.Cells(9).Style.BackColor = Color.Green Then
                bgcolor = "green;"
            End If
            output &= "<td style='" & "background-color:" & bgcolor & "'>"
            output &= grid.Rows(row.Index).Cells(9).Value
            output &= "</td>"

            'end first row of the table
            output &= "</tr>"

        Next

        'end table1
        output &= "</table>"
        output &= "<h4>" & combo.Items(6) & "</h4>"
        output &= "</body>"
        output &= "</html>"
        Return output
    End Function

    Public Function ConvertRangeGridToHtml(ByVal grid As DataGridView, ByVal title As String)
        RTNRSL.ProgressBar1.Visible = True
        RTNRSL.ProgressBar1.PerformStep()
        Dim output As String = "" ' a string to be used as the HTML output
        output &= "<!DOCTYPE HTML PUBLIC -//W3C//DTD HTML 4.0 Transitional//EN>"
        output &= "<html>"
        output &= "<head>"
        output &= "<title>" & title & "</title>"
        output &= "<style type=""text/css""" & ">" & ""
        'create style 1
        output &= ".style1"
        output &= "{"
        output &= "width: 514px;"
        output &= "}"

        'create style 2
        output &= ".style2"
        output &= "{"
        output &= "width: 216px;"
        output &= "font-weight: bold;"
        output &= "}"
        output &= "</style>"

        Dim richtext As New RichTextBox, combo As New ComboBox
        ReadText(richtext, "C:\reportmanager\facilityinfo.txt")
        For Each line In richtext.Lines
            combo.Items.Add(line)
        Next

        output &= "</head>"
        output &= "<body>"
        output &= "<h1>" & combo.Items(0) & "</h1>"
        output &= "<h3>" & combo.Items(1) & "&nbsp" & combo.Items(2) & "&nbsp" & combo.Items(3) & "</h3>"
        output &= "<h3>Tel:" & combo.Items(4) & "&nbsp; Email:" & combo.Items(5) & "</h3>"
        'start table1
        output &= "<table style=" & """" & "width: 100%;""" & " border=" & """1""" & ">"

        'first row of the table
        output &= "<tr>"
        RTNRSL.ProgressBar1.PerformStep()
        'first column of the table
        For i As Integer = 0 To grid.Columns.Count - 1
            If grid.Columns(i).Visible = True Then
                output &= "<td>"
                output &= "<strong>" & grid.Columns(i).HeaderText & "</strong>"
                output &= "</td>"
            End If
        Next
        'end first row of the table
        output &= "</tr>"


        For r As Integer = 0 To grid.Rows.Count - 1
            output &= "<tr>"
            For i As Integer = 0 To grid.Columns.Count - 1
                If grid.Columns(i).Visible = True Then
                    output &= "<td style='" & "background-color:" & FormatCellColor(r, i, grid) & "'>"
                    output &= grid.Rows(r).Cells(i).Value
                    output &= "</td>"
                End If
            Next
            'end row of the table
            output &= "</tr>"
        Next

        'end table1
        output &= "</table>"
        output &= "<h4>" & combo.Items(6) & "</h4>"
        output &= "</body>"
        output &= "</html>"
        Return output
    End Function

    Private Function FormatCellColor(ByVal row As Integer, ByVal cell As Integer, ByVal grid As DataGridView)
        'conditional color formatting
        Dim bgcolor As String = "white;"
        If IsNumeric(grid.Rows(row).Cells(cell).Value) Then
            If grid.Rows(row).Cells(cell).Style.BackColor = Color.Red Then
                bgcolor = "#ff0000;"
            ElseIf grid.Rows(row).Cells(cell).Style.BackColor = Color.Orange Then
                bgcolor = "orange;"
            ElseIf grid.Rows(row).Cells(cell).Style.BackColor = Color.Yellow Then
                bgcolor = "yellow;"
            ElseIf grid.Rows(row).Cells(cell).Style.BackColor = Color.Green Then
                bgcolor = "green;"
            End If
        End If
        Return bgcolor
    End Function
End Class
