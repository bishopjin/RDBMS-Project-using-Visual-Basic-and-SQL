Imports MySql.Data.MySqlClient

Public Class Form9
    Public MysqlConn As MySqlConnection
    Public cmd As New MySqlCommand
    Public da As New MySqlDataAdapter
    Public options As String

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

    Public Sub DisplayAllRecord()
        Dim sql As String
        Dim TempTable As New DataTable

        sql = ""
        Try
            MysqlConnection()
            sql = "SELECT UID, employee_name, timeIn, DATE_FORMAT(Login_date, '%b %d %Y')as LoginDate, " _
                & "timeOut,  DATE_FORMAT(Logout_date, '%b %d %Y')as LogoutDate FROM tbllog;"
            'bind the connection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With

            da.SelectCommand = cmd
            da.Fill(TempTable)

            DataGridView1.DataSource = TempTable
            da.Dispose()
            MysqlConn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub searchRecord()
        Dim sql As String
        Dim num As Double
        Dim TempTable As New DataTable

        sql = ""
        Try
            MysqlConnection()
            If TextBox1.Text = "" Then
                MsgBox("Nothing to search.")
                TextBox1.Focus()
            Else
                If Double.TryParse(TextBox1.Text, num) Then
                    sql = "SELECT UID, employee_name, timeIn, DATE_FORMAT(Login_date, '%b %d %Y')as LoginDate, " _
                        & "timeOut,  DATE_FORMAT(Logout_date, '%b %d %Y')as LogoutDate FROM tbllog " _
                        & "WHERE UID = " & num
                Else
                    sql = "SELECT UID, employee_name, timeIn, DATE_FORMAT(Login_date, '%b %d %Y')as LoginDate, " _
                        & "timeOut,  DATE_FORMAT(Logout_date, '%b %d %Y')as LogoutDate FROM tbllog " _
                        & "WHERE employee_name LIKE '%" & TextBox1.Text & "%' OR timeIn LIKE '%" & TextBox1.Text & "%' OR " _
                        & "Login_date LIKE '%" & TextBox1.Text & "%' OR timeOut LIKE '%" & TextBox1.Text & "%' OR " _
                        & "Logout_date LIKE '%" & TextBox1.Text & "%';"
                End If
                'bind the connection and query
                With cmd
                    .Connection = MysqlConn
                    .CommandText = sql
                End With

                da.SelectCommand = cmd
                da.Fill(TempTable)

                DataGridView1.DataSource = TempTable
                da.Dispose()
                MysqlConn.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim x As String = MsgBox("You sure?", vbYesNo + vbQuestion, "Logout")
        If x = vbYes Then
            Me.Hide()
            Form2.Show()
        Else
            x = vbCancel
        End If
    End Sub

    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        searchRecord()
    End Sub
End Class