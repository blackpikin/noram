Imports System.Data.SqlClient
Imports System.IO
Public Class login
    Dim Authenticate As New Auth
    Dim txtreader As New TextReader
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim connection = New SqlConnection(My.Resources.cstring)
        Dim command = New SqlCommand("SELECT man FROM man", connection)
        Dim reader As SqlDataReader
        Dim installdate As String = ""
        Try
            connection.Open()
            reader = command.ExecuteReader()
            While reader.Read
                installdate = reader.Item(0)
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "Oops! Role not found coz of an error", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            command.Dispose()
            connection.Close()
        End Try

        Dim startdate As Date = CDate(Authenticate.XOrENC(installdate, 215))
        Dim expirydate As Date = CDate(Label5.Text)
        Dim daysleft As Integer = DateDiff("d", startdate, expirydate)

        If installdate = "" Then
            MessageBox.Show("Eish! Its seems that is application has expired. Sorry for any inconvenience")
            Exit Sub
        End If

        If (daysleft <= 0) Or (daysleft > 30) Then
            MessageBox.Show("Eish! Its seems that is application has expired. Sorry for any inconvenience")
            Exit Sub
        Else
            connection = New SqlConnection(My.Resources.cstring)
            command = New SqlCommand("SELECT username, password, role, permissions FROM security WHERE username = '" & TextBox1.Text & "' AND password = '" & Authenticate.XOrENC(TextBox2.Text, 215) & "'", connection)
            Dim username As String = "", pw As String = "", permission = ""
            Try
                connection.Open()
                reader = command.ExecuteReader()
                While reader.Read
                    username = reader.Item(0)
                    pw = reader.Item(1)
                    permission = reader.Item(3)
                End While
            Catch ex As Exception
                MessageBox.Show(ex.Message & vbLf & "Oops! Role not found coz of an error", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                command.Dispose()
                connection.Close()
            End Try
            If username <> "" And pw <> "" Then
                Home.Label2.Text = username
                Home.Label3.Text = pw
                Home.Label4.Text = permission
                Home.Show()
                Me.Close()
            Else
                MessageBox.Show("Invalid username or password")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub login_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not File.Exists("C:\reportmanager\man.txt") Then
            txtreader.WriteText(Date.Now.Date, "C:\reportmanager\man.txt")
            Dim connection = New SqlConnection(My.Resources.cstring)
            Dim command = New SqlCommand("INSERT INTO man (man) VALUES('" & Authenticate.XOrENC(Date.Now.Date, 215) & "')", connection)
            Try
                connection.Open()
                command.ExecuteNonQuery()
            Catch ex As Exception
                MessageBox.Show(ex.Message & vbLf & "There was an error during the save." & vbLf & "Please check the data and try saving again", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                command.Dispose()
                connection.Close()
            End Try
        End If

    End Sub

   
    
    
End Class