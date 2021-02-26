Imports System.IO
Imports MySql.Data.MySqlClient

Public Class Form8
    Public MysqlConn As MySqlConnection
    Public cmd As New MySqlCommand
    Public dr As MySqlDataReader
    Public da As New MySqlDataAdapter

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

    Function getSHA1Hash(ByVal strToHash As String) As String
        Dim sha1Obj As New Security.Cryptography.SHA1CryptoServiceProvider
        Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(strToHash)

        bytesToHash = sha1Obj.ComputeHash(bytesToHash)

        Dim strResult As String = ""

        For Each b As Byte In bytesToHash
            strResult += b.ToString("x2")
        Next

        Return strResult
    End Function

    Function checkAccountExist() As Boolean
        Dim sql, hashPass, utype, user As String
        Dim accountable As New DataTable
        Dim result As Boolean

        Try
            hashPass = getSHA1Hash(TextBox7.Text)
            MysqlConnection()
            sql = "SELECT * FROM tblUser WHERE User ='" & TextBox6.Text & "' AND Pass = '" & hashPass & "' OR " _
                & "Fname = '" & TextBox2.Text & "' AND MName = '" & TextBox3.Text & "' AND " _
                & "LName = '" & TextBox4.Text & "' AND Gender = '" & TextBox5.Text & "';"
            'bind the MysqlConnnection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With

            da.SelectCommand = cmd
            da.Fill(accountable)

            'check if theres a result by getting the count number of rows
            If accountable.Rows.Count > 0 Then
                'it gets the data from specific column and assign to the variable
                utype = accountable.Rows(0).Item(8)
                user = accountable.Rows(0).Item(6)
                MsgBox("Account already exist with username: " & user & " and access level of: " & utype, vbOKOnly + vbInformation, "Account Verification")
                result = True
            Else
                result = False
            End If
            da.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MysqlConn.Close()

        Return result
    End Function

    Function getAccessLevel() As Integer
        Dim sql As String
        Dim level As Integer
        Dim accountable As New DataTable

        Try
            MysqlConnection()
            sql = "SELECT MAX(Type) FROM tblUser"
            'bind the MysqlConnnection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With

            da.SelectCommand = cmd
            da.Fill(accountable)

            'check if theres a result by getting the count number of rows
            If accountable.Rows.Count > 0 And Not accountable.Rows(0).Item(0) Is DBNull.Value Then
                'it gets the data from specific column and assign to the variable
                level = 2
            Else
                level = 1
            End If
            da.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MysqlConn.Close()
        Return level
    End Function

    Private Sub saveNewUser()
        Dim hashPass, accesslevel As String
        Dim day, month, year, result As Integer

        day = DateTimePicker1.Value.Day
        month = DateTimePicker1.Value.Month
        year = DateTimePicker1.Value.Year
        accesslevel = getAccessLevel()

        Try
            hashPass = getSHA1Hash(TextBox7.Text)
            If Not checkAccountExist() Then
                If TextBox7.Text <> TextBox8.Text Then
                    MsgBox("Password dont match.", vbOK + vbCritical, "Invalid")
                    TextBox7.Clear()
                    TextBox8.Clear()
                    TextBox7.Focus()
                Else
                    MysqlConnection()
                    With cmd
                        .Connection = MysqlConn
                        .CommandText = "INSERT INTO tblUser(FName, MName, LName, Gender, BDate, User, Pass, Type) VALUES" _
                            & "('" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "', " _
                            & "'" & TextBox5.Text & "', '" & year & "-" & month & "-" & day & "', " _
                            & "'" & TextBox6.Text & "', '" & hashPass & "', " & TextBox1.Text & ");"

                        result = cmd.ExecuteNonQuery
                        'if the result is equal to zero it means that no rows is inserted or somethings wrong during the execution
                        If result = 0 Then
                            MsgBox("Something went wrong..", vbOKOnly + vbExclamation, "System")
                        Else
                            MsgBox("New Record successfully saved!", vbOKOnly + vbInformation, "System")
                        End If
                    End With
                    MysqlConn.Close() 'closes connection
                    clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Function checkusername()
        Dim sql, user As String
        Dim usernameTable As New DataTable
        sql = "select * from tblUser where User ='" & TextBox6.Text & "'"
        'bind the MysqlConnnection and query

        With cmd
            .Connection = MysqlConn
            .CommandText = sql
        End With
        da.SelectCommand = cmd
        da.Fill(usernameTable)
        If usernameTable.Rows.Count > 0 Then
            user = "exist"
        Else
            user = "none"
        End If
        Return user
    End Function

    Private Sub clear()
        TextBox1.Text = getAccessLevel()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox2.Focus()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim x As String = MsgBox("You sure?", vbYesNo + vbQuestion, "Back")
        If x = vbYes Then
            Me.Hide()
            Form1.Show()
        Else
            x = vbCancel
        End If
        clear()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ask As String = MsgBox("All data input is correct?", vbYesNo + vbQuestion, "Save Data")

        If ask = vbYes Then
            If checkusername() = "none" Then
                saveNewUser()
                Me.Hide()
                Form1.Show()
                Form1.TextBox1.Focus()
            Else
                MsgBox("Username already exist, select another one.", vbOKOnly + vbCritical, "Invalid username")
                TextBox6.SelectAll()
                TextBox6.Focus()
            End If
        Else
            ask = vbCancel
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If TextBox2.Text <> "" And TextBox3.Text <> "" And TextBox4.Text <> "" And TextBox5.Text <> "" And TextBox6.Text <> "" And TextBox7.Text <> "" And TextBox8.Text <> "" Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = getAccessLevel()
        CenterToScreen()
    End Sub
End Class