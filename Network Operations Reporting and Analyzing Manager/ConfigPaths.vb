Imports System.IO
Public Class configPaths
    Dim txtrd As New TextReader
    Dim quote As String = """"
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        txtrd.WriteText(quote & TextBox1.Text & quote, "C:\reportmanager\pathtodump.txt")
        MessageBox.Show("The data was saved")
    End Sub

    Private Sub configPaths_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Icon = Home.Icon
        txtrd.ReadText(TextBox1, "C:\reportmanager\pathtodump.txt")
    End Sub
End Class