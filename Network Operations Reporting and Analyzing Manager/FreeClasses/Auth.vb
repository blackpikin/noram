Imports System.Text
Imports System.Data.SqlClient

Public Class Auth
    Public Function XOrENC(ByVal textToEncrypt As String, ByVal Key As Integer) As String
        Dim inSb As New StringBuilder(textToEncrypt)
        Dim outSb As New StringBuilder(textToEncrypt.Length)
        Dim c As Char
        For i As Integer = 0 To textToEncrypt.Length - 1
            c = inSb(i)
            c = Chr(Asc(c) Xor Key)
            outSb.Append(c)
        Next
        Return outSb.ToString()
    End Function

    Public Function isAuthorized(ByVal permissions As String, ByVal modul As String) As Boolean
        Dim arrModules = Split(permissions, ",")
        Dim combo As New ComboBox
        For Each item In arrModules
            combo.Items.Add(item)
        Next
        If combo.Items.Contains(modul) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function getRole(ByVal role As String) As String
        Dim connection = New SqlConnection(My.Resources.cstring)
        Dim command = New SqlCommand("SELECT role FROM userroles WHERE role = '" & role.ToLower & "'", connection)
        Dim reader As SqlDataReader, existingrole As String = ""
        Try
            connection.Open()
            reader = command.ExecuteReader()
            While reader.Read
                existingrole = reader.Item(0)
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "Oops! Role not found coz of an error", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            command.Dispose()
            connection.Close()
        End Try
        Return existingrole
    End Function

    Public Function getUser(ByVal username As String)
        Dim connection = New SqlConnection(My.Resources.cstring)
        Dim command = New SqlCommand("SELECT username FROM security WHERE username = '" & username & "'", connection)
        Dim reader As SqlDataReader, existinguser As String = ""
        Try
            connection.Open()
            reader = command.ExecuteReader()
            While reader.Read
                existinguser = reader.Item(0)
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "Oops! User not found coz of an error", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            command.Dispose()
            connection.Close()
        End Try
        Return existinguser
    End Function

    Public Sub addNewRole(ByVal role As String)
        If role = getRole(role) Then
            MessageBox.Show("A role with that name already exists")
            Exit Sub
        Else
            Dim connection = New SqlConnection(My.Resources.cstring)
            Dim command = New SqlCommand("INSERT INTO userroles (role) VALUES('" & role & "')", connection)
            Try
                connection.Open()
                command.ExecuteNonQuery()
                MessageBox.Show("New role Added!")
            Catch ex As Exception
                MessageBox.Show(ex.Message & vbLf & "There was an error during the save." & vbLf & "Please check the data and try saving again", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                command.Dispose()
                connection.Close()
            End Try
        End If
    End Sub

    Public Sub deleteRole(ByVal role As String)
        Dim connection = New SqlConnection(My.Resources.cstring)
        Dim command = New SqlCommand("DELETE FROM userroles WHERE role ='" & role & "'", connection)
        Try
            connection.Open()
            command.ExecuteNonQuery()
            MessageBox.Show("Role Deleted!")
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "The operation failed!." & vbLf & "Please check the data and try saving again", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            command.Dispose()
            connection.Close()
        End Try
    End Sub

    Public Sub saveNewUser(ByVal username As String, ByVal password As String, ByVal role As String, ByVal permissions As String)
        If username = getUser(username) Then
            MessageBox.Show("There is already a user with that name")
            Exit Sub
        Else
            Dim connection = New SqlConnection(My.Resources.cstring)
            Dim command = New SqlCommand("INSERT INTO security (username, password, role, permissions) VALUES('" & username & "','" & password & "','" & role & "', '" & permissions & "')", connection)
            Try
                connection.Open()
                command.ExecuteNonQuery()
                MessageBox.Show("New user Added!")
            Catch ex As Exception
                MessageBox.Show(ex.Message & vbLf & "There was an error during the save." & vbLf & "Please check the data and try saving again", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                command.Dispose()
                connection.Close()
            End Try
        End If
    End Sub

    Public Sub deleteUser(ByVal username As String)
        Dim connection = New SqlConnection(My.Resources.cstring)
        Dim command = New SqlCommand("DELETE FROM security WHERE username ='" & username & "'", connection)
        Try
            connection.Open()
            command.ExecuteNonQuery()
            MessageBox.Show("User:" & username & " has been deleted!")
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "The operation failed!." & vbLf & "Please check the data and try saving again", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            command.Dispose()
            connection.Close()
        End Try
    End Sub

    Public Sub deleteAllUsers()
        Dim connection = New SqlConnection(My.Resources.cstring)
        Dim command = New SqlCommand("DELETE FROM security", connection)
        Try
            connection.Open()
            command.ExecuteNonQuery()
            MessageBox.Show("All users have been deleted!")
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "The operation failed!." & vbLf & "Please check the data and try saving again", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            command.Dispose()
            connection.Close()
        End Try
    End Sub

    Public Sub loadAllUsers(ByVal grid As DataGridView)
        Dim connection = New SqlConnection(My.Resources.cstring)
        Dim command = New SqlCommand("SELECT * FROM security", connection)
        Dim reader As SqlDataReader
        grid.Rows.Clear()
        Try
            connection.Open()
            reader = command.ExecuteReader()
            While reader.Read
                Dim row() As String = New String() {reader.Item(1), "", reader.Item(3), reader.Item(4)}
                grid.Rows.Add(row)
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "Oops! Users could not be loaded coz of an error", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            command.Dispose()
            connection.Close()
        End Try
    End Sub
End Class
