Imports MySql.Data.MySqlClient

Public Class Form4
    Public MysqlConn As MySqlConnection
    Public cmd As New MySqlCommand
    Public da As New MySqlDataAdapter
    Public dr As MySqlDataReader
    Dim result, num2 As Integer
    Dim num As Double
    Public state, x As String
    Public user_id As String

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

    Public Sub getValue(objectName As Object, table As String)
        Dim sql As String

        Try
            MysqlConnection()
            If objectName Is TextBox2 Then
                'get the last item ID / row ID in the table
                sql = "SELECT MAX(ShoeID) FROM " & table
            Else
                sql = "SELECT * FROM " & table
            End If

            'bind the connection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With
            dr = cmd.ExecuteReader()
            If objectName Is TextBox2 And dr.Read Then
                If dr.IsDBNull(0) And x = "" Then
                    x = "exit"
                    MsgBox("No existing item in your inventory.", vbOKOnly + vbInformation, "Inventory")
                ElseIf dr.IsDBNull(0) And x = "exit" Then
                    x = ""
                Else
                    'get item ID and add 1 then show in textbox2
                    TextBox2.Text = dr.GetInt32(0) + 1
                End If
            Else
                dr.Dispose()
                dr = cmd.ExecuteReader()
                Do While dr.Read = True
                    objectName.Items.Add(dr.GetString(1))
                Loop
            End If
            MysqlConn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Function getDate() As String
        Dim sql, date_now As String
        Dim publictable As New DataTable

        date_now = ""

        Try
            MysqlConnection()
            sql = "SELECT DATE_FORMAT(now(), '%Y-%m-%e')"
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With
            dr = cmd.ExecuteReader()
            dr.Read()
            date_now = dr.GetString(0)
            dr.Dispose()
            'MysqlConn.Close() 'closes connection
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return date_now
    End Function

    Function checkDuplicateItem()
        Dim result As Boolean
        Dim itemId As Integer
        Dim findtable As New DataTable

        Try
            MysqlConnection()
            With cmd
                .Connection = MysqlConn
                .CommandText = "SELECT ShoeID FROM tblShoe " _
                    & "WHERE (SELECT BID FROM tblBrand WHERE Brand = '" & ComboBox1.Text & "') = BID AND " _
                    & "(SELECT SID FROM tblSize WHERE Size = '" & ComboBox2.Text & "') = SID AND " _
                    & "(SELECT CID FROM tblColor WHERE Color = '" & ComboBox4.Text & "') = CID AND " _
                    & "(SELECT TID FROM tblType WHERE Type = '" & ComboBox3.Text & "') = TID AND " _
                    & "(SELECT CatID FROM tblCategory WHERE Category = '" & ComboBox5.Text & "') = CatID"

                da.SelectCommand = cmd
                da.Fill(findtable)

                'check if theres a result by getting the count number of rows
                If findtable.Rows.Count > 0 Then
                    result = True
                    itemId = findtable.Rows(0).Item(0)
                    MsgBox("Item already exist with item ID: " & itemId, vbOKOnly + vbExclamation, "Record Found")
                Else
                    result = False
                End If
                da.Dispose()
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MysqlConn.Close()

        Return result
    End Function

    Public Sub addstock()
        Dim day, month, year As Integer

        day = DateTimePicker1.Value.Day
        month = DateTimePicker1.Value.Month
        year = DateTimePicker1.Value.Year

        Try
            MysqlConnection()
            With cmd
                .Connection = MysqlConn
                .CommandText = "INSERT INTO tblInventory(ShoeID, Received_Qty, Date_received, Receiver_ID) VALUES " _
                    & "((SELECT max(ShoeID) FROM tblShoe), " & TextBox1.Text & ", '" & year & "-" & month & "-" & day & "'" _
                    & ", " & user_id & ")"

                'in this line it Executes a transact-SQL statements against the connection and returns the number of rows affected 
                result = cmd.ExecuteNonQuery
                'if the result is equal to zero it means that no rows is inserted or somethings wrong during the execution
                If result = 0 Then
                    MsgBox("Something went wrong..", vbOKOnly + vbExclamation, "System")
                Else
                    MsgBox("New Record successfully saved!", vbOKOnly + vbInformation, "System")
                End If
            End With
            MysqlConn.Close() 'closes connection
            TextBox2.Clear()
            TextBox2.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub insertNewItem()
        Try
            If Not checkDuplicateItem() Then
                MysqlConnection()
                With cmd
                    .Connection = MysqlConn
                    .CommandText = "INSERT INTO tblShoe(BID, SID, CID, TID, CatID, Price) VALUES " _
                        & "((SELECT BID FROM tblBrand WHERE Brand = '" & ComboBox1.Text & "'), " _
                        & "(SELECT SID FROM tblSize WHERE Size = " & ComboBox2.Text & "), " _
                        & "(SELECT CID FROM tblColor WHERE Color = '" & ComboBox4.Text & "'), " _
                        & "(SELECT TID FROM tblType WHERE Type = '" & ComboBox3.Text & "'), " _
                        & "(SELECT CatID FROM tblCategory WHERE Category = '" & ComboBox5.Text & "'), " _
                        & TextBox6.Text & ")"

                    'in this line it Executes a transact-SQL statements against the connection and returns the number of rows affected 
                    result = cmd.ExecuteNonQuery
                    'if the result is equal to zero it means that no rows is inserted or somethings wrong during the execution
                    If result = 0 Then
                        MsgBox("Something went wrong..", vbOKOnly + vbExclamation, "System")
                    Else
                        MsgBox("New Record successfully saved!", vbOKOnly + vbInformation, "System")
                    End If
                End With
                MysqlConn.Close() 'closes connection
                getValue(TextBox2, "tblShoe")
                clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub searchItem()
        Dim sql As String
        Dim searchtable As New DataTable

        Try
            MysqlConnection()
            sql = "SELECT tblBrand.Brand, tblSize.Size, tblType.Type, " _
                & "tblColor.Color, tblCategory.Category, tblShoe.Price" _
                & " FROM tblShoe" _
                & " INNER JOIN tblBrand ON tblShoe.BID = tblBrand.BID" _
                & " INNER JOIN tblSize ON tblShoe.SID = tblSize.SID" _
                & " INNER JOIN tblColor ON tblShoe.CID =  tblColor.CID" _
                & " INNER JOIN tblType ON tblShoe.TID = tblType.TID" _
                & " INNER JOIN tblCategory ON tblShoe.CatID = tblCategory.CatID" _
                & " WHERE tblShoe.ShoeID = " & TextBox2.Text & ";"

            'bind the connection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With
            da.SelectCommand = cmd
            da.Fill(searchtable)

            If searchtable.Rows.Count > 0 Then
                ComboBox1.Text = searchtable.Rows(0).Item(0)
                ComboBox2.Text = searchtable.Rows(0).Item(1)
                ComboBox3.Text = searchtable.Rows(0).Item(2)
                ComboBox4.Text = searchtable.Rows(0).Item(3)
                ComboBox5.Text = searchtable.Rows(0).Item(4)
                TextBox6.Text = searchtable.Rows(0).Item(5)
                TextBox1.Focus()
            Else
                TextBox2.Focus()
                MsgBox("Item does not exist.", vbOKOnly + vbCritical, "Not Found")
            End If
            da.Dispose()
            MysqlConn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub clear()
        TextBox2.Enabled = False
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        ComboBox5.Enabled = True
        TextBox6.Enabled = True
        DateTimePicker1.Enabled = True
        'set the default value of the combobox to null
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox5.SelectedIndex = -1
        TextBox1.Clear()
        TextBox6.Clear()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim x As String = MsgBox("You sure?", vbYesNo + vbQuestion, "Back")
        If x = vbYes And state = "add" Then
            getValue(TextBox2, "tblShoe")
            clear()
            Label8.Show()
            Label9.Show()
            TextBox1.Show()
            DateTimePicker1.Show()
            Me.Hide()
            Form2.Show()
        ElseIf x = vbYes And state = "deliver" Then
            clear()
            Button3.Enabled = True
            Me.Hide()
            Form2.Show()
        Else
            x = vbCancel
        End If
    End Sub

    Public Sub loadAllData()
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        getValue(ComboBox1, "tblBrand")
        getValue(ComboBox2, "tblSize")
        getValue(ComboBox3, "tblType")
        getValue(ComboBox4, "tblColor")
        getValue(ComboBox5, "tblCategory")
    End Sub

    Public Sub deliverOption()
        TextBox2.Enabled = True
        TextBox2.Clear()
        TextBox2.Focus()
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        TextBox6.Enabled = False
        DateTimePicker1.Enabled = False
        DateTimePicker1.Text = getDate()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ask As String = MsgBox("All data input is correct?", vbYesNo + vbQuestion, "Save Data")

        If ask = vbYes Then
            If state = "add" Then
                insertNewItem()
            Else
                addstock()
            End If
        Else
            ask = vbCancel
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        Form6.Show()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If state = "add" Then
            If ComboBox1.Text <> "" And ComboBox2.Text <> "" And ComboBox3.Text <> "" And ComboBox4.Text <> "" And ComboBox5.Text <> "" And Double.TryParse(TextBox6.Text, num) Then
                If num > 0.9999 Then
                    Button1.Enabled = True
                Else
                    Button1.Enabled = False
                End If
            Else
                Button1.Enabled = False
            End If
        Else
            If TextBox1.Text <> "" And Integer.TryParse(TextBox1.Text, num2) Then
                If num2 > 0 Then
                    Button1.Enabled = True
                Else
                    Button1.Enabled = False
                End If
            Else
                Button1.Enabled = False
            End If
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If state = "deliver" And TextBox2.TextLength = 5 Then
            searchItem()
        End If
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub
End Class