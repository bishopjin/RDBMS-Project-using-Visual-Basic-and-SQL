Imports MySql.Data.MySqlClient
Imports System.IO

Public Class Form10
    Public MysqlConn As MySqlConnection
    Public cmd As New MySqlCommand
    Dim currenWorkingDirectory, path As String

    Public Sub MysqlConnection()
        MysqlConn = New MySqlConnection()

        'Connection String
        MysqlConn.ConnectionString = "server=localhost;" _
        & "user id=root;" _
        & "password=;" _
        & "database="

        'OPENING THE MysqlConnNECTION
        MysqlConn.Open()

    End Sub

    Private Sub backupDatabase()
        Dim cmdProcess As New Process
        'Dim dateToday As String

        'dateToday = DateTime.Now.ToString("yyyy_MM_dd")
        currenWorkingDirectory = Directory.GetCurrentDirectory()
        path = currenWorkingDirectory & "\backup_database\dbshoeinventory_backup.sql"


        With cmdProcess
            .StartInfo = New ProcessStartInfo("cmd.exe", String.Format("/C C:\xampp\mysql\bin\mysqldump -u root dbshoeinventory > ""{0}"" ", path))
            '.StartInfo = New ProcessStartInfo("cmd.exe", String.Format("/C C:\laragon\bin\mysql\mysql-5.7.24-winx64\bin\mysqldump -u root dbshoeinventory > ""{0}"" ", path))
            With .StartInfo
                .CreateNoWindow = True
                .UseShellExecute = False
                .RedirectStandardOutput = True
            End With
            .Start()
            .WaitForExit()
        End With

        ' Read output to a string variable.
        Dim backupResult As String = cmdProcess.StandardOutput.ReadToEnd

        If backupResult = "" Then
            Label1.Text = ""
            Label2.Text = ""
            MsgBox("Backup of database was created. ", vbOKOnly + vbInformation, "Success")
            Me.Hide()
            Form2.Show()
        Else
            MsgBox(backupResult, vbOKOnly + vbInformation, "Failed")
        End If
    End Sub

    Private Sub restoreDatabase()
        Dim result As Integer
        Dim cmdProcess As New Process
        Dim logresult As Boolean

        currenWorkingDirectory = Directory.GetCurrentDirectory()
        path = currenWorkingDirectory & "\backup_database\dbshoeinventory_backup.sql"

        Try
            MysqlConnection()
            With cmd
                .Connection = MysqlConn
                .CommandText = "CREATE DATABASE dbshoeinventory;"

                result = cmd.ExecuteNonQuery
                'if the result is equal to zero it means that no rows is inserted or somethings wrong during the execution
                If result = 0 Then
                    logresult = False
                Else
                    logresult = True
                End If
                MysqlConn.Close() 'closes connection
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        If logresult Then
            With cmdProcess
                .StartInfo = New ProcessStartInfo("cmd.exe", String.Format("/C C:\xampp\mysql\bin\mysql -u root dbshoeinventory < ""{0}"" ", path))
                '.StartInfo = New ProcessStartInfo("cmd.exe", String.Format("/C C:\laragon\bin\mysql\mysql-5.7.24-winx64\bin\mysql -u root dbshoeinventory < ""{0}"" ", path))
                With .StartInfo
                    .CreateNoWindow = True
                    .UseShellExecute = False
                    .RedirectStandardOutput = True
                End With
                .Start()
                .WaitForExit()
            End With
            ' Read output to a string variable.
            Dim restoreResult As String = cmdProcess.StandardOutput.ReadToEnd

            If restoreResult = "" Then
                Label1.Text = ""
                Label2.Text = ""
                MsgBox("Database restored successfully.", vbOKOnly + vbInformation, "Success")
                Me.Hide()
                Form1.Button3.Enabled = False
                Form1.Button1.Enabled = True
                Form1.setAll()
                Form1.Show()
                'Application.Restart()
            Else
                MsgBox(restoreResult, vbOKOnly + vbInformation, "Failed")
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label1.Text = "Creating a backup of database."
        Label2.Text = "Please wait..."
        backupDatabase()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label1.Text = "Restoring your database."
        Label2.Text = "Please wait..."
        restoreDatabase()
    End Sub

    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub
End Class