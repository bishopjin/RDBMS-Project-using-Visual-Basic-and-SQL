Imports MySql.Data.MySqlClient

Public Class Form7
    Public MysqlConn As MySqlConnection
    Public cmd As New MySqlCommand
    Public da As New MySqlDataAdapter
    Public dr As MySqlDataReader
    Public total_cost As Integer

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
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox12.Clear()
        TextBox2.Focus()
    End Sub

    Public Sub searchItem()
        Dim sql As String
        Dim searchtable As New DataTable

        Try
            MysqlConnection()
            sql = "SELECT tblBrand.Brand, tblSize.Size, tblColor.Color, " _
                & "tblType.Type, tblCategory.Category" _
                & " FROM tblShoe" _
                & " INNER JOIN tblSize ON tblShoe.SID = tblSize.SID" _
                & " INNER JOIN tblColor ON tblShoe.CID =  tblColor.CID" _
                & " INNER JOIN tblType ON tblShoe.TID = tblType.TID" _
                & " INNER JOIN tblCategory ON tblShoe.CatID = tblCategory.CatID" _
                & " INNER JOIN tblBrand ON tblShoe.BID = tblBrand.BID WHERE tblShoe.ShoeID = " & TextBox2.Text & ";"

            'bind the connection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With
            da.SelectCommand = cmd
            da.Fill(searchtable)

            If searchtable.Rows.Count > 0 Then
                TextBox3.Text = searchtable.Rows(0).Item(0)
                TextBox4.Text = searchtable.Rows(0).Item(1)
                TextBox5.Text = searchtable.Rows(0).Item(2)
                TextBox6.Text = searchtable.Rows(0).Item(3)
                TextBox7.Text = searchtable.Rows(0).Item(4)
                TextBox12.Text = getInStockQty()
                TextBox8.Focus()
            Else
                TextBox2.Focus()
                MsgBox("Item does not exist.", vbOKOnly + vbCritical, "Not Found")
            End If
            da.Dispose()
            MysqlConn.Close()
        Catch ex As Exception
            MsgBox("Search: " + ex.Message)
        End Try
    End Sub

    Function generateOrderNumber() As Long
        Dim order_number As Long
        Dim sql As String

        Try
            MysqlConnection()
            sql = "SELECT MAX(orderNumber) FROM tblOrder;"
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With
            dr = cmd.ExecuteReader()
            dr.Read()
            If Not dr.IsDBNull(0) Then
                order_number = Convert.ToInt32(dr.GetString(0)) + 1
            Else
                order_number = 1000000001
            End If
            dr.Dispose()
            MysqlConn.Close() 'closes connection
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return order_number
    End Function

    Function getInStockQty() As Integer
        Dim Total_Received_Qty, Total_Ordered_Qty, in_stock As Integer
        Dim sql, sql2 As String

        Try
            MysqlConnection()
            sql = "SELECT SUM(Received_Qty) FROM tblInventory WHERE tblInventory.ShoeID = " & TextBox2.Text
            sql2 = "SELECT SUM(oQTY) FROM tblOrder WHERE tblOrder.oShoeID = " & TextBox2.Text
            Try
                With cmd
                    .Connection = MysqlConn
                    .CommandText = sql
                End With
                dr = cmd.ExecuteReader()
                dr.Read()
                Total_Received_Qty = dr.GetInt32(0)
                dr.Dispose()
            Catch ex As Exception
                Total_Received_Qty = 0
            End Try

            Try
                With cmd
                    .Connection = MysqlConn
                    .CommandText = sql2
                End With
                dr = cmd.ExecuteReader()
                dr.Read()
                Total_Ordered_Qty = dr.GetInt32(0)
                dr.Dispose()
            Catch ex As Exception
                Total_Ordered_Qty = 0
            End Try
            MysqlConn.Close() 'closes connection

            in_stock = Total_Received_Qty - Total_Ordered_Qty
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return in_stock
    End Function

    Public Sub saveOrder()
        Dim result As Integer

        Try
            MysqlConnection()
            With cmd
                .Connection = MysqlConn
                .CommandText = "INSERT INTO tblOrder(orderNumber, oShoeID, oQTY, oDATE) VALUES " _
                    & "(" & TextBox9.Text & ", " & TextBox2.Text & ", " _
                    & TextBox8.Text & ", Date_Format(now(), '%Y-%m-%e'))"

                'in this line it Executes a transact-SQL statements against the connection and returns the number of rows affected 
                result = cmd.ExecuteNonQuery
                'if the result is equal to zero it means that no rows is inserted or somethings wrong during the execution
                If result = 0 Then
                    MsgBox("Something went wrong..", vbOKOnly + vbExclamation, "Order")
                Else
                    MsgBox("Order successfully saved!", vbOKOnly + vbInformation, "Order")
                End If
            End With
            MysqlConn.Close() 'closes connection
            clear()
            DisplayAllOrder()
            displayTotal()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub displayTotal()
        Dim sql As String
        Try
            MysqlConnection()
            sql = "SELECT SUM(tblOrder.oQTY) FROM tblOrder WHERE tblOrder.orderNumber = " & TextBox9.Text
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With
            dr = cmd.ExecuteReader()
            dr.Read()
            TextBox11.Text = dr.GetInt32(0)
            dr.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DisplayAllOrder()
        Dim sql As String
        Dim TempTable As New DataTable

        Try
            MysqlConnection()
            sql = "SELECT tblOrder.orderNumber as Order_Number, tblShoe.ShoeID as ItemID, tblBrand.Brand, tblSize.Size, tblColor.Color, " _
                & "tblType.Type, tblCategory.Category, tblShoe.Price, tblOrder.oQTY as Order_Quantity, " _
                & "(tblOrder.oQTY * tblShoe.Price) as Total_Price" _
                & " FROM tblOrder" _
                & " INNER JOIN tblShoe ON tblOrder.oShoeID = tblShoe.ShoeID" _
                & " INNER JOIN tblBrand ON tblShoe.BID = tblBrand.BID" _
                & " INNER JOIN tblSize ON tblShoe.SID = tblSize.SID" _
                & " INNER JOIN tblColor ON tblShoe.CID =  tblColor.CID" _
                & " INNER JOIN tblType ON tblShoe.TID = tblType.TID" _
                & " INNER JOIN tblCategory ON tblShoe.CatID = tblCategory.CatID" _
                & " WHERE tblOrder.orderNumber = " & TextBox9.Text & ";"

            'bind the connection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With

            da.SelectCommand = cmd
            da.Fill(TempTable)
            total_cost += (Convert.ToInt32(TempTable.Rows(0).Item(7)) * Convert.ToInt32(TempTable.Rows(0).Item(8)))
            TextBox10.Text = total_cost
            DataGridView1.DataSource = TempTable
            da.Dispose()
            MysqlConn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    'search item
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        searchItem()
    End Sub
    'close
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim x As String = MsgBox("You sure?", vbYesNo + vbQuestion, "Back")
        If x = vbYes Then
            DataGridView1.DataSource = Nothing
            total_cost = 0
            Button1.Enabled = False
            TextBox10.Text = ""
            TextBox11.Text = ""
            clear()
            Me.Hide()
            Form2.Show()
        Else
            x = vbCancel
        End If
    End Sub
    'save order
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim in_stock, num As Integer
        Dim ask As String

        in_stock = getInStockQty()
        Double.TryParse(TextBox8.Text, num)

        If TextBox3.Text = "" Then
            MsgBox("No item selected", vbOKOnly + vbInformation, "None")
            TextBox2.Focus()
        ElseIf in_stock >= Convert.ToInt32(TextBox8.Text) And num > 0 Then
            ask = MsgBox("Save item order?", vbYesNo + vbQuestion, "Save Order")

            If ask = vbYes Then
                saveOrder()
            Else
                ask = vbCancel
            End If
        Else
            MsgBox("There are " & in_stock & " left in the inventory.", vbOK + vbCritical, "Out of Stock")
            TextBox8.Focus()
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.TextLength = 5 Then
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub
End Class