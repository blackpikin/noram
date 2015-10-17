Imports System.Environment
Imports System.IO
Public Class Home
    Dim Authenticate As New Auth
    Private Sub EndSessionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EndSessionToolStripMenuItem.Click
        If MessageBox.Show("Are you trying to leave?", "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            MessageBox.Show("This application will now exit", "Exit", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
            login.ShowDialog()
        Else
        End If

    End Sub

    Private Sub ESToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Hide()
        AlarmMonitoring.Show()
    End Sub

    Private Sub LinkAliToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkAliToolStripMenuItem.Click
        Me.Hide()
        LinkAli.ShowDialog()
    End Sub

    Private Sub RTNCapacityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RTNCapacityToolStripMenuItem.Click
        Me.Hide()
        EtheCapacity.ShowDialog()
    End Sub

    Private Sub RSLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RSLToolStripMenuItem.Click
        If Authenticate.isAuthorized(Label4.Text, "Performance Management") = True Then
            RTNRSL.Show()
            Me.Hide()
        Else
            MessageBox.Show("You do not have permission to view this area")
            Exit Sub
        End If

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub FacilityInfoToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FacilityInfoToolStripMenuItem1.Click
        If Authenticate.isAuthorized(Label4.Text, "Administrator") = True Then
            facility.ShowDialog()
        Else
            MessageBox.Show("You do not have permission to view this")
            Exit Sub
        End If

    End Sub

    Private Sub Home_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Label9.Text = Date.Now.ToString("MMM dd yyyy   hh:mm:ss")

        If Not Directory.Exists("C:\reportmanager") Then
            Directory.CreateDirectory("C:\reportmanager")
        Else
            If Not File.Exists("C:\reportmanager\facilityinfo.txt") Then
                File.Create("C:\reportmanager\facilityinfo.txt")
            End If
            If Not File.Exists("C:\reportmanager\pathtodump.txt") Then
                File.Create("C:\reportmanager\pathtodump.txt")
            End If
        End If
    End Sub

    Private Sub SetUpFilePathsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SetUpFilePathsToolStripMenuItem.Click
        configPaths.ShowDialog()
    End Sub

    Private Sub RegisterUsersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegisterUsersToolStripMenuItem.Click
        If Authenticate.isAuthorized(Label4.Text, "Administrator") = True Then
            manageusers.ShowDialog()
        Else
            MessageBox.Show("You do not have permission to view this")
            Exit Sub
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label9.Text = Date.Now.ToString("MMM dd yyyy   hh:mm:ss")
    End Sub

    Private Sub IPRANUtilizationCheckToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub IPRANUtilizationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPRANUtilizationToolStripMenuItem.Click
        Me.Hide()
        IPRANRTN.Show()
    End Sub
End Class
