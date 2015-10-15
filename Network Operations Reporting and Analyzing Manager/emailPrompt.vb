Imports Microsoft.Office.Interop

Public Class emailPrompt
    Dim txtread As New TextReader
    Public mode As String

    Public Sub sendEmail(ByVal subject As String, ByVal msgBody As String, ByVal reciever As String)
        Try
            Dim oApp As Outlook._Application
            oApp = New Outlook.Application()
            Dim oMsg As Outlook._MailItem
            oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)
            oMsg.Subject = subject
            oMsg.BodyFormat = 2
            oMsg.Body = msgBody
            oMsg.To = reciever
            Dim sBodyLen As String = oMsg.Body.Length
            Dim oAttachs As Outlook.Attachments = oMsg.Attachments
            oMsg.Send()
            oApp = Nothing
            oMsg = Nothing
            oAttachs = Nothing
            MessageBox.Show("Email sent. Click OK to continue")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Please enter a valid email address")
            Exit Sub
        End If

        Dim reciever As String = TextBox1.Text
        Dim subject As String = ""
        If TextBox2.Text = "" Then
            subject = "Network operations report"
        Else
            subject = TextBox2.Text
        End If
        Dim msgBody As String = ""
        If mode = "Single RSL" Then
            msgBody = txtread.convertGridToHtml(RTNRSL.DataGridView1, "RSL Report")
        ElseIf mode = "Range RSL" Then
            msgBody = txtread.ConvertRangeGridToHtml(RTNRSL.DataGridView2, "RSL Report")
        ElseIf mode = "Single Ethernet" Then
            msgBody = txtread.convertGridToHtml(EtheCapacity.DataGridView1, "Ethernet capacity Report")
        ElseIf mode = "Range Ethernet" Then
            msgBody = txtread.ConvertRangeGridToHtml(EtheCapacity.DataGridView2, "Ethernet capacity Report")
        End If

        'send the email
        sendEmail(subject, msgBody, reciever)
    End Sub

    Private Sub emailPrompt_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
End Class