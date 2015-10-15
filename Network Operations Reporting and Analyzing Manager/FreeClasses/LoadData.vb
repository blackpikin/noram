Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.Runtime.CompilerServices

Public Class LoadData
#Region "Fields"

    Private sqlcon As SqlConnection
    Private sqlcom As SqlCommand
    Private reader As SqlDataReader
    Private adaptor As SqlDataAdapter
    Private ds As New DataSet
    Private i As Integer
    Private sql As String
    Private cstring As String = My.Resources.cstring
#End Region

#Region "Properties"

    Public Property ConnectionToSQL
        Get
            Return sqlcon
        End Get
        Set(ByVal value)
            sqlcon = value
        End Set
    End Property

    Public Property CommandForSQL
        Get
            Return sqlcom
        End Get
        Set(ByVal value)
            sqlcom = value
        End Set
    End Property

    Public Property ReaderForSQL
        Get
            Return reader
        End Get
        Set(ByVal value)
            reader = value
        End Set
    End Property

    Public Property AdaptorForSQL
        Get
            Return adaptor
        End Get
        Set(ByVal value)
            adaptor = value
        End Set
    End Property

    Public ReadOnly Property ConnectionString
        Get
            Return cstring
        End Get
    End Property
#End Region

#Region "SQLConnection"

    Public Sub connect(ByVal tb As TextBox)
        sqlcon.ConnectionString = My.Resources.cstring
        sqlcom = New SqlCommand("select * from Alarms", sqlcon)
        sqlcon.Open()
        reader = sqlcom.ExecuteReader
        While reader.Read
            tb.Text = reader.Item(0)
        End While
    End Sub

    Public Sub connect(ByVal dgv As DataGridView)
        sqlcon.ConnectionString = My.Resources.cstring
        sqlcom = New SqlCommand("select * from Alarms", sqlcon)
        sqlcon.Open()
        reader = sqlcom.ExecuteReader
        While reader.Read
            dgv.Text = reader.Item(0)
        End While
    End Sub

    Public Sub connect()
        Dim sqlcon As New SqlConnection
        sqlcon.ConnectionString = My.Resources.cstring
        sqlcom = New SqlCommand("select count from Alarms", sqlcon)
        sqlcon.Open()
        ' tb.Text = reader.Item(0)
    End Sub
#End Region

#Region "Validation"
    'to deny empty control fields'
    Public Sub DenyEmpty(ByVal ctrl As Control)
        If ctrl.Text = "" Then
            MessageBox.Show("You must fill all fields", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
            ctrl.BackColor = Color.Pink
        Else
            ctrl.BackColor = Color.Lime
            Exit Sub
        End If
    End Sub

    'to deny numerical value in control fields'
    Public Sub DenyNumeric(ByVal ctrl As Control)
        If IsNumeric(ctrl.Text) Then
            Beep()
            ctrl.ResetText()
        ElseIf ctrl.Text = "" Then
            ctrl.BackColor = Color.Pink
        Else
            ctrl.BackColor = Color.Lime
        End If
    End Sub

    Public Sub DenyString(ByVal kontrol As Control)
        If (Not IsNumeric(kontrol.Text)) And (kontrol.Text <> "") Then
            Beep()
            kontrol.ResetText()
            MessageBox.Show("Only figures are allowed in this field!", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
            kontrol.BackColor = Color.Pink
        Else
            kontrol.BackColor = Color.Lime
        End If
    End Sub

    Public Sub denyWhite(ByVal kontrol As Control)
        If kontrol.Text = "" Then
            Exit Sub
        End If
    End Sub

    Public Sub denyChoice(ByVal kontrol As Control)
        If kontrol.Text = "Choose one" Then
            kontrol.BackColor = Color.Pink
            MessageBox.Show("You must fill all fields, Choose an item in the list!", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Else
            kontrol.BackColor = Color.Lime
            Exit Sub
        End If
    End Sub

    Public Sub mustContain(ByVal kontrol As ComboBox)
        If Not kontrol.Items.Contains(kontrol.Text) Then
            kontrol.BackColor = Color.Pink
            Exit Sub
        Else
            kontrol.BackColor = Color.Lime
            Exit Sub
        End If
    End Sub

#End Region

#Region "Save Methods"


    Public Sub SaveData(ByVal table As String, ByVal columns As String, ByVal parameters As String)
        sqlcon = New SqlConnection(ConnectionString)
        sqlcom = New SqlCommand("Delete * from RTNRSL")
        sqlcom = New SqlCommand("INSERT INTO " & table & " (" & columns & ") VALUES (" & parameters & ")", sqlcon)

        Try
            sqlcon.Open()
            sqlcom.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "There was an error during the save." & vbLf & "Please check the data and try saving again", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            sqlcom.Dispose()
            sqlcon.Close()

        End Try
        Exit Sub
    End Sub
#End Region

#Region "GetData"
    Public Sub GetData()
        sqlcon = New SqlConnection(ConnectionString)
        sqlcom = New SqlCommand("SELECT * FROM Alarms ", sqlcon)

        Try
            sqlcon.Open()
            reader = sqlcom.ExecuteReader()
            While reader.Read
                Dim row() As String = New String() {reader.Item(0), reader.Item(1), reader.Item(2), reader.Item(3), reader.Item(4), reader.Item(5), reader.Item(6)}
                AlarmMonitoring.DataGridView1.Rows.Add(row)
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbLf & "There was an error during the save." & vbLf & "Please check the data and try saving again", My.Resources.appname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            sqlcom.Dispose()
            sqlcon.Close()
        End Try
    End Sub
#End Region


#Region "Enrolled"
    Public Sub Counter(ByVal tb As TextBox, ByVal table As String, ByVal columns As String)
        sqlcon = New SqlConnection(My.Resources.cstring)
        sqlcom = New SqlCommand("SELECT COUNT (FirstName)  FROM  enroll", sqlcon)
        Try
            tb.Clear()
            sqlcon.Open()

            Dim kount As Integer = (sqlcom.ExecuteScalar())
            tb.Text = Str(kount)
        Catch ex As Exception
            tb.Text = "0"
        Finally
            sqlcon.Close()
        End Try
    End Sub
#End Region

#Region "LoadData"
    Public Sub LoadData(ByVal tb As TextBox, ByVal table As String, ByVal i As Integer, ByVal columns As String)
        sqlcon = New SqlConnection(ConnectionString)
        sqlcom = New SqlCommand("SELECT " & columns & " FROM " & table, sqlcon)
        Try
            tb.Clear()
            sqlcon.Open()
            reader = sqlcom.ExecuteReader()
            While reader.Read
                tb.Text = reader.Item(0)
            End While
        Catch ex As Exception
            tb.Text = "0"
        Finally
            sqlcon.Close()
        End Try
    End Sub

    
    Public Sub LoadDataDataGridView(ByVal grid As DataGridView, ByVal table As String, ByVal i As Integer, ByVal columns As String)

        sqlcon = New SqlConnection(ConnectionString)
        sqlcom = New SqlCommand("SELECT " & columns & " FROM " & table & "", sqlcon)
        Try
            grid.Rows.Clear()
            sqlcon.Open()
            reader = sqlcom.ExecuteReader()
            While reader.Read
                Dim row() As String = New String() {reader.Item(i), reader.Item(i + 1), reader.Item(i + 2), reader.Item(i + 3), reader.Item(i + 4), reader.Item(i + 5), reader.Item(i + 6), reader.Item(i + 7), reader.Item(i + 8), reader.Item(i + 9), reader.Item(i + 10), reader.Item(i + 11), reader.Item(i + 12), reader.Item(i + 13), reader.Item(i + 14), reader.Item(i + 15), reader.Item(i + 16), reader.Item(i + 17), reader.Item(i + 18), reader.Item(i + 19), reader.Item(i + 20), reader.Item(i + 21), reader.Item(i + 22), reader.Item(i + 23), reader.Item(i + 24), reader.Item(i + 25), reader.Item(i + 26), reader.Item(i + 27), reader.Item(i + 28), reader.Item(i + 29), reader.Item(i + 30), reader.Item(i + 31), reader.Item(i + 32), reader.Item(i + 33), reader.Item(i + 34), reader.Item(i + 35), reader.Item(i + 36), reader.Item(i + 37), reader.Item(i + 38), reader.Item(i + 39), reader.Item(i + 40), reader.Item(i + 41), reader.Item(i + 42), reader.Item(i + 43), reader.Item(i + 44), reader.Item(i + 45), reader.Item(i + 46), reader.Item(i + 47), reader.Item(i + 48)}
                grid.Rows.Add(row)
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            sqlcon.Close()
        End Try
    End Sub
    Public Function DefaultIfDBNull(Of T)(ByVal obj As Object) As T
        Return If(Convert.IsDBNull(obj), CType(Nothing, T), CType(obj, T))
    End Function
    Sub LoadData(ByVal grid As DataGridView, ByVal table As String, ByVal columns As String, ByVal i As Integer)
        sqlcon = New SqlConnection(ConnectionString)
        sqlcom = New SqlCommand("SELECT " & columns & " FROM " & table & "", sqlcon)
        Try
            grid.Rows.Clear()
            sqlcon.Open()
            reader = sqlcom.ExecuteReader()
            While reader.Read
                '  If IsDBNull(ds.Tables(0).Rows(i).Item(0)) Then
                Dim row() As String = New String() {reader.Item(i), reader.Item(i + 1).ToString, reader.Item(i + 2), reader.Item(i + 3), reader.Item(i + 4), reader.Item(i + 5), reader.Item(i + 6), reader.Item(i + 7), reader.Item(i + 8), reader.Item(i + 9), reader.Item(i + 10), reader.Item(i + 11), reader.Item(i + 12), reader.Item(i + 13), reader.Item(i + 14), reader.Item(i + 15), reader.Item(i + 16), reader.Item(i + 17), reader.Item(i + 18), reader.Item(i + 19), reader.Item(i + 20), reader.Item(i + 21), reader.Item(i + 22), reader.Item(i + 23), reader.Item(i + 24), reader.Item(i + 25), reader.Item(i + 26), reader.Item(i + 27), reader.Item(i + 28), reader.Item(i + 29), reader.Item(i + 30), reader.Item(i + 31), reader.Item(i + 32), reader.Item(i + 33), reader.Item(i + 34), reader.Item(i + 35), reader.Item(i + 36), reader.Item(i + 37), reader.Item(i + 38), reader.Item(i + 39), reader.Item(i + 40), reader.Item(i + 41), reader.Item(i + 42), reader.Item(i + 43), reader.Item(i + 44), reader.Item(i + 45), reader.Item(i + 46), reader.Item(i + 47), reader.Item(i + 48)}
                grid.Rows.Add(row)
                '  End If

            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            sqlcon.Close()
        End Try
    End Sub

    Public Sub LoadData(ByVal list As ListBox, ByVal table As String, ByVal i As Integer, ByVal columns As String)
        sqlcon = New SqlConnection(ConnectionString)
        sqlcom = New SqlCommand("SELECT " & columns & " FROM " & table & "", sqlcon)
        Try
            sqlcon.Open()
            reader = sqlcom.ExecuteReader()
            list.Items.Clear()
            While reader.Read
                list.Items.Add(reader.Item(i))
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            sqlcon.Close()
        End Try
    End Sub
    Public Sub LoadData(ByVal checklist As CheckedListBox, ByVal table As String, ByVal i As Integer, ByVal columns As String)
        sqlcon = New SqlConnection(ConnectionString)
        sqlcom = New SqlCommand("SELECT " & columns & " FROM " & table & "", sqlcon)
        Try
            sqlcon.Open()
            reader = sqlcom.ExecuteReader()
            checklist.Items.Clear()
            While reader.Read
                checklist.Items.Add(reader.Item(i))
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            sqlcon.Close()
        End Try
    End Sub
    Public Sub LoadData(ByVal rtb As RichTextBox, ByVal table As String, ByVal i As Integer, ByVal columns As String)
        sqlcon = New SqlConnection(ConnectionString)
        sqlcom = New SqlCommand("SELECT " & columns & " FROM " & table & "", sqlcon)
        Try
            sqlcon.Open()
            reader = sqlcom.ExecuteReader()
            rtb.ResetText()
            While reader.Read
                rtb.Text = reader.Item(i)
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            sqlcon.Close()
        End Try
    End Sub


#End Region

#Region "Delete Methods"
    Public Sub DeleteData(ByVal table As String, ByVal col As String, ByVal col2 As String, ByVal wher As String, ByVal wher2 As String)
        sqlcon = New SqlConnection(ConnectionString)
        sqlcom = New SqlCommand("DELETE FROM " & table & " WHERE " & col & " = '" & wher & "' AND " & col2 & " = '" & wher2 & "'", sqlcon)
        Try
            sqlcon.Open()
            sqlcom.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            sqlcon.Close()
        End Try
    End Sub
    Public Sub DeleteData(ByVal table As String, ByVal column As String)
        sqlcon = New SqlConnection(ConnectionString)
        sqlcom = New SqlCommand("DELETE FROM " & table & " ", sqlcon)
        Try
            sqlcon.Open()
            sqlcom.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            sqlcon.Close()
        End Try
    End Sub
#End Region

End Class
