Imports System.IO

Public Class Welcome


    Public Function Readfile(ByVal ThaFile As String)
        Dim strcontents As String
        Dim sr As StreamReader
        sr = New StreamReader(ThaFile)
        strcontents = sr.ReadToEnd()
        sr.Close()
        Return strcontents
    End Function
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        login.Show()
        Me.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Timer1.Enabled = True
    End Sub

End Class