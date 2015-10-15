

Imports System.Data.SqlClient
Imports System.Data.DataTable
Imports System.IO

Public Class AlarmMonitoring

    Dim LD As New LoadData
    Dim sda As SqlDataAdapter
    Dim sqlcon As New SqlConnection
    Dim connection As New SqlConnection(My.Resources.cstring)
    Dim sqlQuery As String
    Dim counter As Integer

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Me.Hide()
        Home.Show()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ComboBox1.SelectedItem = "Choose one" Then
            MessageBox.Show("You must select an Alarm!!!", "Morning Drill", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Exit Sub
        End If


        Dim Datetime As DateTime
        Dim getdate As Date
        Datetime = getdate
        Dim str As String = My.Resources.cstring
        Dim sqlcon As New SqlConnection(str)
        Dim sqlcom As New SqlCommand("select * from RTNAlarms  where Severity =  '" & ComboBox1.SelectedItem & "'", sqlcon) '[Last Occurred ]  between '" & DateTimePicker1.Text & "' and '" & DateTimePicker2.Text & "' and 
        Dim reader As SqlDataReader

        Try
            DataGridView1.Rows.Clear()
            sqlcon.Open()
            reader = sqlcom.ExecuteReader()
            While reader.Read
                Dim row() As String = New String() {reader.Item(0), reader.Item(0 + 1), reader.Item(0 + 2), reader.Item(0 + 3), reader.Item(0 + 4), reader.Item(0 + 5), reader.Item(0 + 6)}
                DataGridView1.Rows.Add(row)

            End While
        Catch err As Exception
            MsgBox(err.Message)

        Finally
            sqlcon.Close()
        End Try
        Exit Sub

    End Sub


    'sqlcom = New SqlCommand("SELECT COUNT (Name) from RTNAlarms WHERE Severity = '" & ComboBox1.Text & "' ", sqlcon)

    '  Try
    ' sqlcon.Open()
    ' reader = (sqlcom.ExecuteReader())
    'While reader.Read
    'If ComboBox1.Text = "Critical" Then
    ' TextBox3.Clear()
    ' TextBox4.Clear()
    ' TextBox2.Text = reader
    ' ElseIf ComboBox1.Text = "Major" Then
    ' TextBox2.Clear()
    ' TextBox4.Clear()
    '   TextBox3.Text = reader
    '  ElseIf ComboBox1.Text = "Minor" Then
    'TextBox2.Clear()
    ' TextBox3.Clear()
    '  TextBox4.Text = reader
    ' Else
    '     reader = "0"
    ' End If

    '  End While
    '  Catch err As Exception
    '  MsgBox(Err.Message)
    ' Finally
    ' sqlcon.Close()
    ' End Try

    Private Sub AlarmMonitoring_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label7.Text = Date.Now.ToString("MMM dd yyyy   hh:mm:ss")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView1.Rows.Clear()
    End Sub


    Private Sub DeleteSelectedItemToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteSelectedItemToolStripMenuItem.Click
        Dim connection As New SqlConnection(My.Resources.cstring)
        Dim sqlcom As New SqlCommand("Delete Name from CurrentAlarms", connection)
        Dim reader As SqlDataReader

        Try
            sqlcon.Open()
            reader = sqlcom.ExecuteReader()
            While reader.Read
                Dim row() As String = New String() {reader.Item(0), reader.Item(0 + 1), reader.Item(0 + 2), reader.Item(0 + 3), reader.Item(0 + 4), reader.Item(0 + 5), reader.Item(0 + 6)}
                DataGridView1.Rows.Clear()

            End While
        Catch err As Exception
            MessageBox.Show("Alarm Cleared!!!", "Morning Drill", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Finally
            sqlcon.Close()
        End Try
    End Sub

End Class
