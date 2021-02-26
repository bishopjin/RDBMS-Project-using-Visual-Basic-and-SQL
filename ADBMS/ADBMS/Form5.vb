Imports MySql.Data.MySqlClient

Public Class Form5
    Public MysqlConn As MySqlConnection
    Public cmd As New MySqlCommand
    Public da As New MySqlDataAdapter
    Public result As Integer
    Public num As Double

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

    Public Sub UpdateRecord()
        Try
            MysqlConnection()
            With cmd
                .Connection = MysqlConn
                .CommandText = "UPDATE tblShoe set tblShoe.Price = " & TextBox6.Text & " WHERE ShoeID = " & TextBox7.Text & ";"

                'in this line it Executes a transact-SQL statements against the connection and returns the number of rows affected 
                result = cmd.ExecuteNonQuery

                'if the result is equal to zero it means that no rows is inserted or somethings wrong during the execution
                If result = 0 Then
                    MsgBox("Something went wrong..", vbOKOnly + vbExclamation, "System")
                Else
                    MsgBox("Record successfully updated!", vbOKOnly + vbInformation, "System")
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MysqlConn.Close() 'closes connection
    End Sub
    Private Sub clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim x As String = MsgBox("You sure?", vbYesNo + vbQuestion, "Exit and discard change")
        If x = vbYes Then
            clear()
            Form3.DataGridView1.Refresh()
            Me.Hide()
            Form3.Show()
        Else
            x = vbCancel
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ask As String = MsgBox("Data input is correct?", vbYesNo + vbQuestion, "Update Data")

        If ask = vbYes Then
            UpdateRecord()
            clear()
            Me.Hide()
            Form3.DisplayAllRecord("inventory")
            Form3.Show()
        Else
            ask = vbCancel
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox3.Text <> "" And TextBox4.Text <> "" And TextBox5.Text <> "" And TextBox6.Text <> "" And TextBox7.Text <> "" Then
        'If TextBox6.Text <> "" And IsNumeric(TextBox6.Text) And TextBox6.Text <> "0" Then
        If Double.TryParse(TextBox6.Text, num) Then
            If num > 0.9999 Then
                Button2.Enabled = True
            Else
                Button2.Enabled = False
            End If

        Else
            Button2.Enabled = False
        End If
    End Sub

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox6.Focus()
        CenterToScreen()
    End Sub
End Class