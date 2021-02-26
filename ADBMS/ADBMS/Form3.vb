Imports MySql.Data.MySqlClient

Public Class Form3
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

    Public Sub DisplayAllRecord(typeofrecord As String)
        Dim sql As String
        Dim TempTable As New DataTable

        sql = ""
        Try
            MysqlConnection()
            If typeofrecord = "inventory" Then
                sql = "SELECT tblShoe.ShoeID as ItemID, tblBrand.Brand, tblSize.Size, tblColor.Color, " _
                    & "tblType.Type, tblCategory.Category, tblShoe.Price, tblInventory.Deliver_Qty as Delivered_Quantity, " _
                    & "tblInventory.date_deliver as Delivered_Date, tblInventory.Stock_qty as Total_Stock" _
                    & " FROM tblShoe" _
                    & " INNER JOIN tblBrand ON tblShoe.BID = tblBrand.BID" _
                    & " INNER JOIN tblSize ON tblShoe.SID = tblSize.SID" _
                    & " INNER JOIN tblColor ON tblShoe.CID =  tblColor.CID" _
                    & " INNER JOIN tblType ON tblShoe.TID = tblType.TID" _
                    & " INNER JOIN tblCategory ON tblShoe.CatID = tblCategory.CatID" _
                    & " INNER JOIN tblInventory ON tblShoe.ShoeID = tblInventory.ShoeID;"
            Else
                sql = "SELECT tblOrder.orderNumber as Order_Number, tblShoe.ShoeID as ItemID, tblBrand.Brand, tblSize.Size, tblColor.Color, " _
                    & "tblType.Type, tblCategory.Category, tblShoe.Price, tblOrder.oQTY as Order_Quantity, " _
                    & "(tblOrder.oQTY * tblShoe.Price) as Total_Price" _
                    & " FROM tblOrder" _
                    & " INNER JOIN tblShoe ON tblOrder.oShoeID = tblShoe.ShoeID" _
                    & " INNER JOIN tblBrand ON tblShoe.BID = tblBrand.BID" _
                    & " INNER JOIN tblSize ON tblShoe.SID = tblSize.SID" _
                    & " INNER JOIN tblColor ON tblShoe.CID =  tblColor.CID" _
                    & " INNER JOIN tblType ON tblShoe.TID = tblType.TID" _
                    & " INNER JOIN tblCategory ON tblShoe.CatID = tblCategory.CatID;"
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
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub searchRecord()
        Dim sql As String
        Dim num As Integer
        Dim oNum As Long
        Dim searchtable As New DataTable

        sql = ""
        Try
            MysqlConnection()
            If options = "inventory" Then
                sql = "SELECT tblShoe.ShoeID as ItemID, tblBrand.Brand, tblSize.Size, tblColor.Color, " _
                    & "tblType.Type, tblCategory.Category, tblShoe.Price, tblInventory.Deliver_Qty as Delivered_Quantity, " _
                    & "tblInventory.date_deliver as Delivered_Date, tblInventory.Stock_qty as Total_Stock" _
                    & " FROM tblShoe" _
                    & " INNER JOIN tblBrand ON tblShoe.BID = tblBrand.BID" _
                    & " INNER JOIN tblSize ON tblShoe.SID = tblSize.SID" _
                    & " INNER JOIN tblColor ON tblShoe.CID =  tblColor.CID" _
                    & " INNER JOIN tblType ON tblShoe.TID = tblType.TID" _
                    & " INNER JOIN tblCategory ON tblShoe.CatID = tblCategory.CatID" _
                    & " INNER JOIN tblInventory ON tblShoe.ShoeID = tblInventory.ShoeID WHERE tblBrand.Brand= '" & TextBox1.Text & "';"
            Else
                If TextBox1.Text <> "" And Integer.TryParse(TextBox1.Text, num) Then
                    If num > 0 Then
                        oNum = num
                    Else
                        oNum = 1000000000
                    End If
                End If
                sql = "SELECT tblOrder.orderNumber as Order_Number, tblShoe.ShoeID as ItemID, tblBrand.Brand, tblSize.Size, tblColor.Color, " _
                    & "tblType.Type, tblCategory.Category, tblShoe.Price, tblOrder.oQTY as Order_Quantity, " _
                    & "(tblOrder.oQTY * tblShoe.Price) as Total_Price" _
                    & " FROM tblOrder" _
                    & " INNER JOIN tblShoe ON tblOrder.oShoeID = tblShoe.ShoeID" _
                    & " INNER JOIN tblBrand ON tblShoe.BID = tblBrand.BID" _
                    & " INNER JOIN tblSize ON tblShoe.SID = tblSize.SID" _
                    & " INNER JOIN tblColor ON tblShoe.CID =  tblColor.CID" _
                    & " INNER JOIN tblType ON tblShoe.TID = tblType.TID" _
                    & " INNER JOIN tblCategory ON tblShoe.CatID = tblCategory.CatID WHERE orderNumber = " & oNum & ";"
            End If
            'bind the connection and query
            With cmd
                .Connection = MysqlConn
                .CommandText = sql
            End With

            da.SelectCommand = cmd
            da.Fill(searchtable)

            If searchtable.Rows.Count > 0 Then
                MsgBox("Record found!", vbOKOnly + vbInformation, "System")
                TextBox1.Focus()
                DataGridView1.DataSource = searchtable
            Else
                TextBox1.Focus()
                MsgBox("Record not found.", vbOKOnly + vbCritical, "System")
            End If
            da.Dispose()
            MysqlConn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        TextBox1.Clear()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        searchRecord()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim x As String = MsgBox("You sure?", vbYesNo + vbQuestion, "Back")
        If x = vbYes Then
            TextBox1.Clear()
            TextBox1.Enabled = False
            Button1.Enabled = False
            Me.Hide()
            Form2.Show()
        Else
            x = vbCancel
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If options = "inventory" Then
            Form5.TextBox1.Text = DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
            Form5.TextBox2.Text = DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value
            Form5.TextBox3.Text = DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value
            Form5.TextBox4.Text = DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value
            Form5.TextBox5.Text = DataGridView1.Item(5, DataGridView1.CurrentRow.Index).Value
            Form5.TextBox6.Text = DataGridView1.Item(6, DataGridView1.CurrentRow.Index).Value
            Form5.TextBox7.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            Me.Hide()
            Form5.Show()
        End If
        
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        options = "inventory"
        DisplayAllRecord("inventory")
        Button1.Enabled = True
        TextBox1.Enabled = True
        TextBox1.Text = "Type the brand name"
        TextBox1.Focus()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        options = "order"
        DisplayAllRecord("order")
        Button1.Enabled = True
        TextBox1.Enabled = True
        TextBox1.Text = "Type the order number"
        TextBox1.Focus()
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub
End Class