Imports System.Data.SqlClient

Public Class manageusers
    Dim Authenticate As New Auth
    Private Sub manageusers_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Icon = Home.Icon
        LoadRoles()
        Authenticate.loadAllUsers(DataGridView1)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Authenticate.addNewRole(TextBox3.Text.ToLower)
        TextBox3.Text = ""
        LoadRoles()
    End Sub

    Private Sub LoadRoles()
        ComboBox1.Items.Clear()
        Dim connection = New SqlConnection(My.Resources.cstring)
        Dim command = New SqlCommand("SELECT role FROM userroles", connection)
        Dim reader As SqlDataReader
        Try
            connection.Open()
            reader = command.ExecuteReader()
            While reader.Read
                ComboBox1.Items.Add(reader.Item(0))
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "Oops! Role not found coz of an error", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            command.Dispose()
            connection.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
            Dim permissions As String = ""
            For Each item In CheckedListBox1.CheckedItems
                permissions = permissions & item & ","
            Next
            Authenticate.saveNewUser(TextBox1.Text, Authenticate.XOrENC(TextBox2.Text, 215), ComboBox1.Text, permissions)
            TextBox1.ResetText()
            TextBox2.ResetText()
            ComboBox1.ResetText()
        CheckedListBox1.ClearSelected()
        Authenticate.loadAllUsers(DataGridView1)
    End Sub

    Private Sub DeleteThisUserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteThisUserToolStripMenuItem.Click
        If MessageBox.Show("Do you really want to delete this user?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Authenticate.deleteUser(DataGridView1.CurrentRow.Cells(0).Value)
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        Else
            Exit Sub
        End If
    End Sub

    Private Sub DeleteAllUsersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteAllUsersToolStripMenuItem.Click
        If MessageBox.Show("Do you really want to delete all users?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Authenticate.deleteAllUsers()
            DataGridView1.Rows.Clear()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub DeleteThisRoleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteThisRoleToolStripMenuItem.Click
        If ComboBox1.Text = "" Then
            MessageBox.Show("Select a role to delete")
            Exit Sub
        End If

        If MessageBox.Show("Do you really want to delete this role?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Authenticate.deleteRole(ComboBox1.Text)
            ComboBox1.Items.Remove(ComboBox1.Text)
            ComboBox1.Text = ""
        Else
            Exit Sub
        End If
    End Sub
End Class