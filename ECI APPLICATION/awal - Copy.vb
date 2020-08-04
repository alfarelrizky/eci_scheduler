Imports System.Data.Odbc
Public Class awal
    Dim currenttime As String
    Dim messagetime As String
    Dim am(50) As String
    Dim pm(50) As String
    Dim connection As OdbcConnection
    Dim reader As OdbcDataReader
    Dim adapter As OdbcDataAdapter
    Dim dataset As New DataSet
    Dim cmd As OdbcCommand
    Dim str As String
    Sub koneksi()
        str = "dsn=prog_1;server=localhost;uid=root;pwd=;database=prog_1"
        'str = "dsn=prog_2;server=localhost;uid=root;pwd=;database=pis_2"
        connection = New OdbcConnection(str)
        If connection.State = ConnectionState.Closed Then
            Try
                connection.Open()
            Catch ex As Exception
                MsgBox(ex.Message)
                MsgBox("database bermasalah", MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        currenttime = TimeOfDay.ToString("hh:mm:ss tt")
        Label1.Text = currenttime
    End Sub
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        am(1) = "00:00:00 AM"
        am(2) = "02:00:00 AM"
        am(3) = "04:00:00 AM"
        am(4) = "06:00:00 AM"
        am(5) = "08:00:00 AM"
        am(6) = "10:00:00 AM"
        am(7) = "12:00:00 AM"
        am(8) = "14:00:00 AM"
        am(9) = "16:00:00 AM"
        am(10) = "18:00:00 AM"
        am(11) = "20:00:00 AM"
        am(12) = "22:00:00 AM"

        pm(1) = "00:00:00 PM"
        pm(2) = "02:00:00 PM"
        pm(3) = "04:00:00 PM"
        pm(4) = "06:00:00 PM"
        pm(5) = "08:00:00 PM"
        pm(6) = "10:00:00 PM"
        pm(7) = "12:00:00 PM"
        pm(8) = "14:00:00 PM"
        pm(9) = "16:00:00 PM"
        pm(10) = "18:00:00 PM"
        pm(11) = "20:00:00 PM"
        pm(12) = "22:00:00 PM"
        messagetime = "08:53:10 AM"
        Label4.Text = "remind set for : " + messagetime
        If ((currenttime = am(1)) Or (currenttime = am(2)) Or (currenttime = am(3)) Or (currenttime = am(4)) Or (currenttime = am(5)) Or (currenttime = am(6)) Or (currenttime = am(7)) Or (currenttime = am(8)) Or (currenttime = am(9)) Or (currenttime = am(10)) Or (currenttime = am(11)) Or (currenttime = am(12)) Or (currenttime = pm(1)) Or (currenttime = pm(2)) Or (currenttime = pm(3)) Or (currenttime = pm(4)) Or (currenttime = pm(5)) Or (currenttime = pm(6)) Or (currenttime = pm(7)) Or (currenttime = pm(8)) Or (currenttime = pm(9)) Or (currenttime = pm(10)) Or (currenttime = pm(11)) Or (currenttime = pm(12))) Then
            'Timer2.Stop()
            'MsgBox(TextBox2.Text)
            Dim tampil As String = "SELECT *from eci_app where keterangan = 'belum'"
            cmd = New OdbcCommand(tampil, connection)
            reader = cmd.ExecuteReader
            If reader.Read() Then
                notifikasi.outputnoeci.Text = ""
                notifikasi.outputnumberpart.Text = ""
                notifikasi.outputpartname.Text = ""
                notifikasi.outputactualdate.Text = ""

                notifikasi.outputnoeci.Text = reader("no_eci")
                notifikasi.outputnumberpart.Text = reader("number_part")
                notifikasi.outputpartname.Text = reader("part_name")
                notifikasi.outputactualdate.Text = reader("actual")
                notifikasi.Show()
                Me.Hide()
                Timer2.Stop()
            End If
            'Button1.Enabled = True
            'Button2.Enabled = False
            'Label4.Text = ""
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Timer2.Start()
        Button1.Enabled = False
        Button2.Enabled = True
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Timer2.Stop()
        Button1.Enabled = True
        Button2.Enabled = False
        Label4.Text = ""
    End Sub
    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
        NotifyIcon1.Visible = False
    End Sub
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        NotifyIcon1.Visible = True
        Me.Hide()
        NotifyIcon1.ShowBalloonTip(3000)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        NotifyIcon1.Visible = True
        NotifyIcon1.ShowBalloonTip(3000)
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub awal_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call koneksi()
        Timer2.Start()
        Button1.Enabled = False
        Button2.Enabled = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub
End Class