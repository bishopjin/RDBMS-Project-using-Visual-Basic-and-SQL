Imports MySql.Data.MySqlClient
Imports System.IO

Public Class Form1
    Public MysqlConn As MySqlConnection
    Public cmd As New MySqlCommand
    Public da As New MySqlDataAdapter
    Public Userid As Integer
    Public userfullname As String
    Public dr As MySqlDataReader

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

    Function checkConnection() As Boolean
        Dim con As Boolean
        Try
            MysqlConnection()
            con = True
            MysqlConn.Close()
        Catch ex As Exception
            con = False
        End Try
        Return con
    End Function

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

    Function logUser()
        Dim result As Integer
        Dim logresult, sql, sql2, timein, datein As String
        Dim publictable As New DataTable

        logresult = ""
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

            timein = publictable.Rows(0).Item(0)
            datein = publictable.Rows(0).Item(1)

            With cmd
                .Connection = MysqlConn
                .CommandText = "INSERT INTO tbllog(UID, employee_name, timeIn, Login_date) VALUES" _
                    & "(" & Userid & ", '" & userfullname & "','" & timein & "' ,'" & datein & "');"

                result = cmd.ExecuteNonQuery
                'if the result is equal to zero it means that no rows is inserted or somethings wrong during the execution
                If result = 0 Then
                    logresult = "failed"
                Else
                    logresult = "success"
                End If
            End With

            sql2 = "SELECT MAX(LogID) FROM tbllog WHERE UID = " & Userid
            'bind the MysqlConnnection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql2
            End With
            dr = cmd.ExecuteReader()
            dr.Read()
            Form2.form2_logID = dr.GetInt32(0)

            MysqlConn.Close() 'closes connection
        Catch ex As Exception
            MsgBox("LOG FAILED: " + ex.Message)
        End Try

        Return logresult
    End Function

    Public Sub checkUser()
        Dim sql, sql2, hashPass As String
        Dim usernameTable, publictable As New DataTable
        Dim utype, uId, lastname, firstname, middleinitial, state As String

        Try
            'check if the textbox is equal to nothing then it will display the message below!.
            If TextBox1.Text = "" And TextBox2.Text = "" Then
                MsgBox("Please enter username", vbOKOnly + vbCritical, "Account Verification")
                TextBox1.Focus()
            ElseIf TextBox2.Text = "" Then
                MsgBox("Please enter password", vbOKOnly + vbCritical, "Account Verification")
                TextBox2.Focus()
            Else
                hashPass = getSHA1Hash(TextBox2.Text)
                MysqlConnection()
                'sql = "select * from tblLogin where User ='" & TextBox1.Text & "' and Pass = '" & hashPass & "'"
                sql = "select * from tblUser where User ='" & TextBox1.Text & "'"
                'bind the MysqlConnnection and query

                With cmd
                    .Connection = MysqlConn
                    .CommandText = sql
                End With
                da.SelectCommand = cmd
                da.Fill(usernameTable)

                'check if theres user exist
                If usernameTable.Rows.Count > 0 Then
                    state = "user"
                Else
                    state = "no_user"
                    MsgBox("Account not valid. Please contact your administrator for assistance", vbOKOnly + vbCritical, "Account Verification")
                    clear()
                End If

                If state = "user" Then
                    sql2 = "select * from tblUser where User ='" & TextBox1.Text & "' and Pass = '" & hashPass & "'"
                    'bind the MysqlConnnection and query

                    With cmd
                        .Connection = MysqlConn
                        .CommandText = sql2
                    End With
                    da.SelectCommand = cmd
                    da.Fill(publictable)

                    If publictable.Rows.Count > 0 Then
                        'it gets the data from specific column and assign to the variable
                        firstname = publictable.Rows(0).Item(1)
                        middleinitial = publictable.Rows(0).Item(2)
                        lastname = publictable.Rows(0).Item(3)
                        utype = publictable.Rows(0).Item(8)
                        'user = publictable.Rows(0).Item(6)
                        uId = publictable.Rows(0).Item(0)
                        'user id
                        Userid = Convert.ToInt32(uId)
                        userfullname = firstname + " " + middleinitial + " " + lastname
                        'check if the type of user is admin
                        If utype = "1" Then
                            'welcomes the user as Admiistrator
                            If logUser() = "success" Then
                                MsgBox("Welcome " & vbNewLine & userfullname, vbOKOnly + vbInformation, "Account Verification")
                                clear()
                                Form7.TextBox1.Text = uId
                                Form4.user_id = uId
                                Form2.Show()
                                Me.Hide()
                            End If
                        Else
                            If logUser() = "success" Then
                                MsgBox("Welcome " & vbNewLine & userfullname, vbOKOnly + vbInformation, "Account Verification")
                                clear()
                                'hide the button - for admin only
                                Form2.Button1.Hide()
                                Form2.Button5.Hide()
                                'resize the button
                                Form2.Button3.Size = New Size(363, 58)
                                Form2.Button4.Size = New Size(363, 58)
                                Form7.TextBox1.Text = uId
                                Form4.user_id = uId
                                Form2.Button6.Hide()
                                Form2.Show()
                                Me.Hide()
                            End If
                        End If
                        'full name of user
                        Form2.label_name.Text = firstname + " " + middleinitial + " " + lastname
                        Form2.label_id.Text = uId
                    Else
                        MsgBox("Invalid password.", vbOKOnly + vbCritical, "Account Verification")
                        TextBox2.SelectAll()
                    End If
                    da.Dispose()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MysqlConn.Close()
    End Sub

    Private Sub clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.Focus()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        checkUser()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim x As String = MsgBox("You sure?", vbYesNo + vbQuestion, "Close Application")
        If x = vbYes Then
            clear()
            Application.Exit()
        Else
            x = vbCancel
        End If
    End Sub

    Private Sub getTimeAndDate()
        Dim sql As String
        Dim publictable As New DataTable

        Try
            MysqlConnection()
            sql = "SELECT DATE_FORMAT(now(), '%W'), DATE_FORMAT(now(), '%M %e'), DATE_FORMAT(now(), '%Y'), TIME_FORMAT(now(), '%h')," _
                & "TIME_FORMAT(now(), '%i'), TIME_FORMAT(now(), '%s'), TIME_FORMAT(now(), '%p')"
            'bind the MysqlConnnection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With

            da.SelectCommand = cmd
            da.Fill(publictable)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Label4.Text = publictable.Rows(0).Item(0)
        Label5.Text = publictable.Rows(0).Item(3) + ":" + publictable.Rows(0).Item(4) + ":" + publictable.Rows(0).Item(5) + " " +
                        publictable.Rows(0).Item(6)
        Label6.Text = publictable.Rows(0).Item(1)
        Label7.Text = publictable.Rows(0).Item(2)
    End Sub

    Public Sub setAll()
        If checkConnection() Then
            Timer1.Enabled = True
            CenterToScreen()
            getTimeAndDate()
            Button3.Hide()
        Else
            MsgBox("Error in database connection. Check your mysql connection.")
            'Application.Exit()
            CenterToScreen()
            Button1.Enabled = False
            Button3.Show()
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        getTimeAndDate()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        setAll()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.Hide()
        Form8.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        Form10.Label1.Text = ""
        Form10.Label2.Text = ""
        Form10.Button1.Enabled = False
        Form10.Show()
    End Sub
End Class
