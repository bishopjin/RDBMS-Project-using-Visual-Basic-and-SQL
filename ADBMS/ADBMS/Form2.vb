Imports MySql.Data.MySqlClient

Public Class Form2
    Public MysqlConn As MySqlConnection
    Public cmd As New MySqlCommand
    Public da As New MySqlDataAdapter
    Public form2_logID As Integer

    Public Sub MysqlConnection()
        MysqlConn = New MySqlConnection()

        'Connection String
        MysqlConn.ConnectionString = "server=localhost;" _
        & "user id=root;" _
        & "password=;" _
        & "database=dbShoeInventory"

        'OPENING THE MysqlConnNECTION
        MysqlConn.Open()

    End Sub

    Private Sub logoutLog()
        Dim result As Integer
        Dim sql, timeout, dateout As String
        Dim publictable As New DataTable

        Try
            MysqlConnection()
            sql = "SELECT TIME_FORMAT(now(), '%h:%i:%s'), DATE_FORMAT(now(), '%Y-%m-%e')"
            'bind the MysqlConnnection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With

            da.SelectCommand = cmd
            da.Fill(publictable)

            timeout = publictable.Rows(0).Item(0)
            dateout = publictable.Rows(0).Item(1)

            With cmd
                .Connection = MysqlConn
                .CommandText = "UPDATE tbllog SET timeOut = '" & timeout & "', Logout_date = '" & dateout & "' WHERE LogID = " & form2_logID
                result = cmd.ExecuteNonQuery
            End With
            MysqlConn.Close() 'closes connection
        Catch ex As Exception
            MsgBox("LOGOUT FAILED: " + ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim x As String = MsgBox("You sure?", vbYesNo + vbQuestion, "Logout")
        If x = vbYes Then
            'Check if hidden
            If Button1.Visible = False Then
                Button1.Show()
                Button5.Show()
                'return to original position
                Button3.Size = New Size(161, 58)
                Button4.Size = New Size(161, 58)
            End If
            If Button6.Visible = False Then
                Button6.Show()
            End If
            'log the date and time
            logoutLog()
            Me.Hide()
            Form1.Show()
        Else
            x = vbCancel
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Form11.Show()
        Form4.state = "add"
        Form4.GroupBox1.Text = "Add New Product"
        Form4.Label8.Hide()
        Form4.Label9.Hide()
        Form4.TextBox1.Hide()
        Form4.DateTimePicker1.Hide()
        Form4.loadAllData()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
        Form7.TextBox9.Text = Form7.generateOrderNumber()
        Form7.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        Form4.state = "deliver"
        Form4.GroupBox1.Text = "Product Deliver"
        Form4.TextBox2.Clear()
        Form4.TextBox2.Focus()
        Form4.Button3.Enabled = False
        Form4.deliverOption()
        Form4.Show()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Hide()
        Form9.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Hide()
        Form10.Label1.Text = ""
        Form10.Label2.Text = ""
        Form10.Button1.Enabled = True
        Form10.Button2.Enabled = False
        Form10.Show()
    End Sub
End Class