Public Class emailSend

    Private Sub emailSend_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Icon = Home.Icon
        WebBrowser1.Navigate("C:\reportmanager\emailpage.htm")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Label6.Visible = True
        Dim tr As New TextReader
        MailHelper.SendMailMessage(TextBox1.Text, TextBox2.Text, TextBox4.Text, TextBox3.Text, TextBox5.Text, tr.convertGridToHtml(RTNRSL.DataGridView1, "RSL Report"))
        Label6.Visible = False
    End Sub
End Class