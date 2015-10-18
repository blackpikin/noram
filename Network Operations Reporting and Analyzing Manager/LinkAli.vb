Imports System.Data.SqlClient

Public Class LinkAli
    Dim LD As New LoadData
    Private sqlcon As SqlConnection
    Private sqlcom As SqlCommand
    Private reader As SqlDataReader
    Private adaptor As SqlDataAdapter
    Private ds As New DataSet
    Private i As Integer
    Private sql As String
    Private cstring As String = My.Resources.cstring

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
        Home.Show()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        LD.DenyString(TextBox1)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox1.BackColor = Color.Lime
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        DataGridView1.Visible = False
        GroupBox2.Visible = True


        Dim conn As New SqlConnection(cstring)
        Dim cmd As New SqlCommand
        Dim dr As SqlDataReader

        ' Open connection, and retrieve dataset.
        conn.Open()

        ' Define Command object
        If RadioButton1.Checked Then
            cmd.CommandText = "Select * From LinkBudget where [Link ID] = '" & TextBox1.Text & "'"
        ElseIf RadioButton2.Checked Then
            cmd.CommandText = "Select * From LinkBudget where Region = '" & ComboBox1.SelectedItem & "'"
        Else
            cmd.CommandText = "Select * From LinkBudget"
        End If

        cmd.CommandType = CommandType.Text
        cmd.Connection = conn

        ' Retrieve data reader.
        dr = cmd.ExecuteReader()

        Dim fieldCount As Integer = dr.FieldCount
        Dim fieldValues(fieldCount - 1) As Object
        Dim headers(fieldCount - 1) As String

        ' Get names of fields. 
        For ctr As Integer = 0 To fieldCount - 1
            headers(ctr) = dr.GetName(ctr)
        Next

        ' Set up data grid.
        DataGridView1.ColumnCount = fieldCount
        ' Get data, replace missing values with "N/A", and display it. 
        Do While dr.Read()
            dr.GetValues(fieldValues)

            For fieldCounter As Integer = 0 To fieldCount - 1
                If Convert.IsDBNull(fieldValues(fieldCounter)) Then
                    fieldValues(fieldCounter) = "N/A"
                End If
            Next
            DataGridView1.Rows.Add(fieldValues)
        Loop
        dr.Close()
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label9.Text = Date.Now.ToString("MMM dd yyyy   hh:mm:ss")
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            Label1.Enabled = True
        Else
            Label1.Enabled = False
        End If
        If RadioButton1.Checked Then
            TextBox1.Enabled = True
        Else
            TextBox1.Enabled = False
            TextBox1.Clear()
            TextBox1.BackColor = Color.SlateGray
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            Label2.Enabled = True
        Else
            Label2.Enabled = False
        End If
        If RadioButton2.Checked Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DataGridView1.Rows.Clear()
    End Sub

    
End Class