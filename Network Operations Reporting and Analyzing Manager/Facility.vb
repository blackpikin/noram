Public Class facility
    Dim txtrd As New TextReader
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        txtrd.WriteText(TextBox1.Text, "C:\reportmanager\facilityinfo.txt")
        txtrd.AppendText(TextBox2, "C:\reportmanager\facilityinfo.txt")
        txtrd.AppendText(TextBox3, "C:\reportmanager\facilityinfo.txt")
        txtrd.AppendText(TextBox4, "C:\reportmanager\facilityinfo.txt")
        txtrd.AppendText(TextBox5, "C:\reportmanager\facilityinfo.txt")
        txtrd.AppendText(TextBox6, "C:\reportmanager\facilityinfo.txt")
        txtrd.AppendText(TextBox7, "C:\reportmanager\facilityinfo.txt")
        MessageBox.Show("The data was saved")
    End Sub

    Private Sub facility_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Icon = Home.Icon
        Dim richtext As New RichTextBox, combo As New ComboBox
        txtrd.ReadText(richtext, "C:\reportmanager\facilityinfo.txt")
        For Each line In richtext.Lines
            combo.Items.Add(line)
        Next

        Try
            TextBox1.Text = combo.Items(0)
            TextBox2.Text = combo.Items(1)
            TextBox3.Text = combo.Items(2)
            TextBox4.Text = combo.Items(3)
            TextBox5.Text = combo.Items(4)
            TextBox6.Text = combo.Items(5)
            TextBox7.Text = combo.Items(6)
        Catch ex As Exception

        End Try
    End Sub
End Class