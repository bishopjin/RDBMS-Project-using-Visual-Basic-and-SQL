Imports MySql.Data.MySqlClient

Public Class Form6
    Public MysqlConn As MySqlConnection
    Public cmd As New MySqlCommand
    Public dr As MySqlDataReader
    Dim result As Integer
    Dim num As Double

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

    Private Sub clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub addNewItemDescription(tableName As String)
        Dim sql, check As String
        Dim tbox As Object

        check = "notfound"

        Try
            MysqlConnection()
            If tableName = "brand" Then
                sql = "SELECT Brand FROM tblBrand"
                tbox = TextBox1
            ElseIf tableName = "size" Then
                sql = "SELECT Size FROM tblSize"
                tbox = TextBox2
            ElseIf tableName = "type" Then
                sql = "SELECT Type FROM tblType"
                tbox = TextBox3
            Else
                sql = "SELECT Color FROM tblColor"
                tbox = TextBox4
            End If

            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With
            dr = cmd.ExecuteReader()
            Do While dr.Read() = True
                If tbox.Text <> "" And dr.GetString(0) = tbox.Text And tbox Is TextBox2 Then
                    check = "found"
                    Exit Do
                ElseIf tbox.Text <> "" And dr.GetString(0).ToLower() = tbox.Text.ToLower() Then
                    check = "found"
                    Exit Do
                End If
            Loop
            dr.Dispose()

            If check = "notfound" Then
                With cmd
                    .Connection = MysqlConn
                    If tableName = "brand" Then
                        .CommandText = "INSERT INTO tblbrand(Brand) VALUES ('" & TextBox1.Text & "');"
                    ElseIf tableName = "size" Then
                        .CommandText = "INSERT INTO tblsize(Size) VALUES ('" & TextBox2.Text & "');"
                    ElseIf tableName = "type" Then
                        .CommandText = "INSERT INTO tbltype(Type) VALUES ('" & TextBox3.Text & "');"
                    Else
                        .CommandText = "INSERT INTO tblcolor(Color) VALUES ('" & TextBox4.Text & "');"
                    End If

                    'in this line it Executes a transact-SQL statements against the connection and returns the number of rows affected 
                    result = cmd.ExecuteNonQuery
                    'if the result is equal to zero it means that no rows is inserted or somethings wrong during the execution
                    If result = 0 Then
                        MsgBox("Something went wrong..", vbOKOnly + vbExclamation, "System")
                    Else
                        MsgBox("New Record successfully saved!", vbOKOnly + vbInformation, "System")
                    End If
                End With
                clear()
            Else
                MsgBox("Value already exist.", vbOK + vbInformation, "Found")
                tbox.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        MysqlConn.Close() 'closes connection
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim x As String = MsgBox("You sure?", vbYesNo + vbQuestion, "Back")
        If x = vbYes Then
            clear()
            Me.Hide()
            Form4.loadAllData()
            Form4.Show()
        Else
            x = vbCancel
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        addNewItemDescription("brand")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        addNewItemDescription("size")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        addNewItemDescription("type")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        addNewItemDescription("color")
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If TextBox1.Text <> "" Then
            Button1.Enabled = True
        ElseIf TextBox2.Text <> "" And Double.TryParse(TextBox2.Text, num) Then
            If num > 0.9999 Then
                Button2.Enabled = True
            Else
                Button2.Enabled = False
            End If
        ElseIf TextBox3.Text <> "" Then
            Button3.Enabled = True
        ElseIf TextBox4.Text <> "" Then
            Button4.Enabled = True
        Else
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
        End If
    End Sub

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub
End Class