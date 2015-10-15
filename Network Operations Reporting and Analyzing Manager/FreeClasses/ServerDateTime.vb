
Imports System.Data

Imports System.Data.SqlClient

Public Class ServerDateTime
    Public dtServerDateTime As DateTime

    Private Shared tmrLocalTimer As New System.Windows.Forms.Timer()

    ' Flag to check whether it is First time run. 

    Dim IsFirstTime As Boolean = True

    Public Sub TimerEventProcessor(ByVal myObject As Object, ByVal myEventArgs As EventArgs)

        'Check whether it is First time run. 

        ' If yes then Get Server time and display it in the Label.

        If IsFirstTime = True Then
            Dim df As String = MorningDrill.My.Resources.cstring

            Dim sqlcomm As SqlCommand

            Dim sqlcon As New SqlConnection

            Dim strConnectionString As String

            Try

                ' Read ConnectionString from Application Configuration file.

                strConnectionString = MorningDrill.My.Resources.cstring

                sqlcon = New SqlConnection(strConnectionString)

                ' Open the Connection

                sqlcon.Open()

                ' Initialize the Command object with commandText and SQLConnection

                sqlcomm = New SqlCommand("Select GetDate()", sqlcon)

                'Execute the query and return the Server Date Time

                dtServerDateTime = sqlcomm.ExecuteScalar

                ' Display the Server Date Time in the Label.

                '  Home.Label.Text = Format(dtServerDateTime, "MM-dd-yyyy")
                ' Form1.Label6.Text = Format(dtServerDateTime, "M-yyyy")
                ' Form1.Label7.Text = Format(dtServerDateTime, "yyyy")
                ' Form1.ToolStripStatusLabel1.Text = "Connected"
                ' Form1.ToolStripStatusLabel1.ForeColor = Color.Green
                ' Set the Flag to False

                IsFirstTime = False

            Catch ex As Exception
                '  Form1.ToolStripStatusLabel1.Text = "Disconnected"
                ' Form1.ToolStripStatusLabel1.ForeColor = Color.Red
                'MsgBox(ex.Message)

            Finally

                If sqlcon.State = ConnectionState.Open Then

                    sqlcomm = Nothing

                    sqlcon.Close()

                End If

            End Try

        Else

            'Add one Second to the dtServerDateTime

            dtServerDateTime = DateAdd(DateInterval.Second, 1, dtServerDateTime)

        End If

        tmrLocalTimer.Enabled = True

        ' Display the Server Date Time in the ToolStripStatusLabel of the MainForm

        ' Form1.Label8.Text = Format(dtServerDateTime, "hh:mm:ss")
        ' Form1.Label5.Text = Format(dtServerDateTime, "MM-dd-yyyy")
        ' Form1.Label6.Text = Format(dtServerDateTime, "M-yyyy")
        ' Form1.Label7.Text = Format(dtServerDateTime, "yyyy")

    End Sub
    Public Sub GetServerDateTime()

        AddHandler tmrLocalTimer.Tick, AddressOf TimerEventProcessor

        ' Sets the timer interval to 1 second.

        tmrLocalTimer.Interval = 1000

        tmrLocalTimer.Start()

    End Sub

End Class
