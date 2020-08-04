Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Data.Odbc
Public Class awal
    Dim currenttime As String
    Dim messagetime As String
    Dim am(50) As String
    Dim pm(50) As String
    Dim n(50) As String
    Dim reader, reader1 As OdbcDataReader
    Dim adapter As OdbcDataAdapter
    Dim cmd, cmd1 As OdbcCommand
    Public jumlah_data As Integer = 0
    Dim str As String
    Public gambarlama(100), gambarbaru(100) As String
    Public hasiltgl As Integer
    Public konfirmasi As String
    Public selisih(100) As Integer
    Public no_eci(100), line(100), model(100), record_nik(100), number_part(100), rack_address(100), jobnumber_new(100), jobnumber_old(100), number_part_baru(100), part_name(100), actual(100), gambar(100), gambar_baru(100), actual_date(100) As String
    Public pointer As Integer

    Function onstartup() As Boolean
        onstartup = (My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", My.Application.Info.AssemblyName, "") = Application.ExecutablePath)
        If onstartup() Then
            My.Computer.Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", My.Application.Info.AssemblyName, Application.ExecutablePath)
            'public registery
        End If
    End Function
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        currenttime = TimeOfDay.ToString("hh:mm:ss tt")
        Label1.Text = currenttime
    End Sub
    Sub hitung()
        DateTimePicker1.Value = System.DateTime.Now.ToString("yyyy/MM/dd")
        Dim sample_tanggal_dari As String = Format(DateAdd(DateInterval.Day, 2, DateTimePicker1.Value), "yyyy-MM-dd")
        Dim akumulasi As Integer = 0
        Dim num_job As Integer = 0
        Dim a As Date
        If nama_pengguna = "" Then
        ElseIf hak = "admin" Then
            jumlah_data = 0
            pointer = 0
            koneksi()
            If connection_stat = True Then
                Try
                    Dim tam As String = "SELECT *from eci_app where ((planning between '" & System.DateTime.Now.ToString("yyyy-MM-dd") & "' and '" & sample_tanggal_dari & "') or planning < '" & System.DateTime.Now.ToString("yyyy-MM-dd") & "') and keterangan = 'belum' AND colom_manggil = '' ORDER BY PLANNING desc"
                    cmd = New OdbcCommand(tam, connection)
                    reader = cmd.ExecuteReader
                    cmd1 = New OdbcCommand(tam, connection)
                    reader1 = cmd1.ExecuteReader
                    If reader1.Read() Then
                        While reader.Read()
                            jumlah_data = jumlah_data + 1
                            no_eci(jumlah_data) = reader("no_eci")
                            rack_address(jumlah_data) = reader("no_ars")
                            line(jumlah_data) = reader("line")
                            model(jumlah_data) = reader("model")
                            number_part(jumlah_data) = reader("number_part")
                            number_part_baru(jumlah_data) = reader("number_part_baru")
                            part_name(jumlah_data) = reader("part_name")
                            actual_date(jumlah_data) = reader("planning")
                            jobnumber_old(jumlah_data) = reader("jobnumber_before")
                            jobnumber_new(jumlah_data) = reader("jobnumber_after")
                            a = actual_date(jumlah_data)
                            selisih(jumlah_data) = DateDiff(DateInterval.Day, a, Today())
                            If reader("gambar_lama") = "" Then
                                gambar(jumlah_data) = "\\10.59.5.200\eci-image\noimage.png"
                            ElseIf reader("gambar_lama") <> "" Then
                                gambar(jumlah_data) = "\\10.59.5.200\eci-image\" & reader("gambar_lama")
                                gambarlama(jumlah_data) = ""
                                gambarlama(jumlah_data) = CStr(reader("gambar_lama"))
                            End If
                            If reader("gambar_baru") = "" Then
                                gambar_baru(jumlah_data) = "\\10.59.5.200\eci-image\noimage.png"
                            ElseIf reader("gambar_baru") <> "" Then
                                gambar_baru(jumlah_data) = "\\10.59.5.200\eci-image\" & reader("gambar_baru")
                                gambarbaru(jumlah_data) = ""
                                gambarbaru(jumlah_data) = CStr(reader("gambar_baru"))
                            End If
                        End While
                    Else

                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        Else
            jumlah_data = 0
            pointer = 0
            koneksi()
            If connection_stat = True Then
                Try
                    Dim use As String
                    Dim tampil2 As String = "SELECT *from eci_app_human where nama = '" & nama_pengguna & "' limit 1"
                    cmd = New OdbcCommand(tampil2, connection)
                    reader = cmd.ExecuteReader
                    While reader.Read()
                        use = reader("zone")
                    End While
                    Dim tampil As String = "SELECT *from eci_app where ((planning between '" & System.DateTime.Now.ToString("yyyy-MM-dd") & "' and '" & sample_tanggal_dari & "') or planning <= '" & System.DateTime.Now.ToString("yyyy-MM-dd") & "') and keterangan = 'belum' AND colom_manggil = ''  AND  line = '" & use & "' ORDER BY PLANNING desc"
                    cmd = New OdbcCommand(tampil, connection)
                    reader = cmd.ExecuteReader
                    cmd1 = New OdbcCommand(tampil, connection)
                    reader1 = cmd1.ExecuteReader
                    If reader1.Read() Then
                        While reader.Read()
                            jumlah_data = jumlah_data + 1
                            no_eci(jumlah_data) = reader("no_eci")
                            rack_address(jumlah_data) = reader("no_ars")
                            line(jumlah_data) = reader("line")
                            model(jumlah_data) = reader("model")
                            number_part(jumlah_data) = reader("number_part")
                            number_part_baru(jumlah_data) = reader("number_part_baru")
                            part_name(jumlah_data) = reader("part_name")
                            actual_date(jumlah_data) = reader("planning")
                            jobnumber_old(jumlah_data) = reader("jobnumber_before")
                            jobnumber_new(jumlah_data) = reader("jobnumber_after")
                            a = actual_date(jumlah_data)
                            selisih(jumlah_data) = DateDiff(DateInterval.Day, a, Today())
                            If reader("gambar_lama") = "" Then
                                gambar(jumlah_data) = "\\10.59.5.200\eci-image\noimage.png"
                            ElseIf reader("gambar_lama") <> "" Then
                                gambar(jumlah_data) = "\\10.59.5.200\eci-image\" & reader("gambar_lama")
                                gambarlama(jumlah_data) = ""
                                gambarlama(jumlah_data) = CStr(reader("gambar_lama"))
                            End If
                            If reader("gambar_baru") = "" Then
                                gambar_baru(jumlah_data) = "\\10.59.5.200\eci-image\noimage.png"
                            ElseIf reader("gambar_baru") <> "" Then
                                gambar_baru(jumlah_data) = "\\10.59.5.200\eci-image\" & reader("gambar_baru")
                                gambarbaru(jumlah_data) = ""
                                gambarbaru(jumlah_data) = CStr(reader("gambar_baru"))
                            End If
                        End While
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        End If
        If jumlah_data <= 0 Then
            jumlah_data = 0
            pointer = 0
            Timer2.Enabled = True
        Else
            pointer += 1
            notifikasi.outputnoeci.Text = ""
            notifikasi.outputnumberpart.Text = ""
            notifikasi.Button3.Text = ""
            notifikasi.Button9.Text = ""
            notifikasi.Button7.Text = ""
            notifikasi.Button12.Text = ""
            notifikasi.Button14.Text = ""
            notifikasi.Button11.Text = ""
            notifikasi.outputpartname.Text = ""
            notifikasi.outputactualdate.Text = ""
            notifikasi.PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
            notifikasi.PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")

            notifikasi.outputnoeci.Text = no_eci(pointer)
            notifikasi.Button14.Text = rack_address(pointer)
            notifikasi.Button11.Text = line(pointer)
            notifikasi.Button9.Text = model(pointer)
            notifikasi.outputnumberpart.Text = number_part(pointer)
            notifikasi.Button3.Text = number_part_baru(pointer)
            notifikasi.outputpartname.Text = part_name(pointer)
            notifikasi.Button12.Text = jobnumber_old(pointer)
            notifikasi.Button7.Text = jobnumber_new(pointer)
            notifikasi.outputactualdate.Text = actual_date(pointer)
            Try
                notifikasi.PictureBox3.Image = Image.FromFile(gambar(pointer))
            Catch ex As Exception
                notifikasi.PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
            End Try
            Try
                notifikasi.PictureBox4.Image = Image.FromFile(gambar_baru(pointer))
            Catch ex As Exception
                notifikasi.PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
            End Try
            If selisih(pointer) > 0 Then
                notifikasi.dedline.Text = "H +" & selisih(pointer)
                notifikasi.dedline.BackColor = Color.Red
                notifikasi.dedline.ForeColor = Color.White
            ElseIf selisih(pointer) = 0 Then
                notifikasi.dedline.Text = "SEKARANG TERAKHIR"
                notifikasi.dedline.BackColor = Color.Red
                notifikasi.dedline.ForeColor = Color.White
            ElseIf selisih(pointer) < 0 Then
                notifikasi.dedline.Text = "H " & selisih(pointer)
                notifikasi.dedline.BackColor = Color.Red
                notifikasi.dedline.ForeColor = Color.White
            Else
                notifikasi.dedline.Text = selisih(pointer) & " HARI"
            End If
            Me.Hide()
            notifikasi.Show()
        End If
    End Sub
    Sub agenda()
        If nama_pengguna = "" Then
        ElseIf hak = "admin" Then
            notifikasi.dedline.BackColor = Color.White
            notifikasi.dedline.BackColor = Color.Green
            Do While jumlah_data > 0
                Dim sudah As String = ""
                Dim tampil As String = "Select *from eci_app where ((keterangan = 'belum') AND (colom_manggil = '')) ORDER BY ACTUAL desc limit 1"
                cmd = New OdbcCommand(tampil, connection)
                reader = cmd.ExecuteReader
                While reader.Read
                    notifikasi.outputnoeci.Text = ""
                    notifikasi.outputnumberpart.Text = ""
                    notifikasi.Button3.Text = ""
                    notifikasi.Button9.Text = ""
                    notifikasi.Button11.Text = ""
                    notifikasi.outputpartname.Text = ""
                    notifikasi.outputactualdate.Text = ""
                    notifikasi.PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
                    notifikasi.PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")

                    notifikasi.outputnoeci.Text = reader("no_eci")
                    notifikasi.Button11.Text = reader("line")
                    notifikasi.Button9.Text = reader("model")
                    notifikasi.outputnumberpart.Text = reader("number_part")
                    notifikasi.Button3.Text = reader("number_part_baru")
                    notifikasi.outputpartname.Text = reader("part_name")
                    notifikasi.outputactualdate.Text = reader("actual")
                    If reader("gambar_lama") = "" Then
                        Return
                    ElseIf reader("gambar_lama") <> "" Then
                    End If
                    If reader("gambar_baru") = "" Then
                        Return
                    ElseIf reader("gambar_baru") <> "" Then
                        notifikasi.PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\" & reader("gambar_baru"))
                    End If
                    'Call pengurangan_waktu(CDate(reader("actual")))
                    If hasiltgl > 0 Then
                        Timer2.Enabled = True
                        notifikasi.dedline.Text = "H -" & hasiltgl
                        notifikasi.dedline.BackColor = Color.Red
                        notifikasi.dedline.ForeColor = Color.White
                    ElseIf hasiltgl = 0 Then
                        Timer2.Enabled = True
                        notifikasi.dedline.Text = "SEKARANG TERAKHIR"
                        notifikasi.dedline.BackColor = Color.Red
                        notifikasi.dedline.ForeColor = Color.White
                    ElseIf hasiltgl < 0 Then
                        Timer2.Enabled = True
                        notifikasi.dedline.Text = "H +" & hasiltgl
                        notifikasi.dedline.BackColor = Color.Red
                        notifikasi.dedline.ForeColor = Color.White
                    Else
                        notifikasi.dedline.Text = hasiltgl & " HARI"
                    End If

                    'release
                    If jumlah_data <= 0 Then
                        notifikasi.Timer1.Stop()
                        Timer2.Enabled = True
                        If connection_stat = True Then
                            Try
                                cmd = connection.CreateCommand
                                cmd.CommandText = "update eci_app set colom_manggil = '' where keterangan = 'belum'"
                                cmd.ExecuteNonQuery()
                            Catch ex As Exception

                            End Try
                        End If
                    Else
                        If connection_stat = True Then
                            Try
                                cmd = connection.CreateCommand
                                cmd.CommandText = "update eci_app set colom_manggil = 'tampil' where no_eci='" & notifikasi.outputnoeci.Text & "' AND number_part = '" & notifikasi.outputnumberpart.Text & "' and number_part_baru='" & notifikasi.Button3.Text & "'"
                                cmd.ExecuteNonQuery()
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                End While
            Loop
        ElseIf hak = "user" Then
            notifikasi.dedline.BackColor = Color.White
            notifikasi.dedline.BackColor = Color.Green
            If jumlah_data = 0 Then
                jumlah_data = 0
                notifikasi.Timer1.Stop()
                Timer2.Enabled = True
                If connection_stat = True Then
                    Try
                        cmd = connection.CreateCommand
                        cmd.CommandText = "update eci_app set colom_manggil = '' where keterangan = 'belum' AND destination = '" & nama_pengguna & "'"
                        cmd.ExecuteNonQuery()
                    Catch ex As Exception

                    End Try
                Else
                    MsgBox("JARINGAN PROBLEM")
                End If
            Else
                notifikasi.Timer1.Enabled = True
                jumlah_data = jumlah_data - 1
                Call koneksi()
                If connection_stat = True Then
                    Dim tampil As String = "SELECT *from eci_app where ((keterangan = 'belum') AND (colom_manggil = '') AND (destination = '" & nama_pengguna & "')) ORDER BY ACTUAL desc"
                    cmd = New OdbcCommand(tampil, connection)
                    reader = cmd.ExecuteReader
                    While reader.Read
                        notifikasi.outputnoeci.Text = ""
                        notifikasi.outputnumberpart.Text = ""
                        notifikasi.Button3.Text = ""
                        notifikasi.Button9.Text = ""
                        notifikasi.Button11.Text = ""
                        notifikasi.outputpartname.Text = ""
                        notifikasi.outputactualdate.Text = ""
                        notifikasi.PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
                        notifikasi.PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")

                        notifikasi.outputnoeci.Text = reader("no_eci")
                        notifikasi.Button11.Text = reader("line")
                        notifikasi.Button9.Text = reader("model")
                        notifikasi.outputnumberpart.Text = reader("number_part")
                        notifikasi.Button3.Text = reader("number_part_baru")
                        notifikasi.outputpartname.Text = reader("part_name")
                        notifikasi.outputactualdate.Text = reader("actual")
                        If reader("gambar_lama") = "" Then
                            Return
                        ElseIf reader("gambar_lama") <> "" Then
                            notifikasi.PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\" & reader("gambar_lama"))
                        End If
                        If reader("gambar_baru") = "" Then
                            Return
                        ElseIf reader("gambar_baru") <> "" Then
                            notifikasi.PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\" & reader("gambar_baru"))
                        End If
                        'Call pengurangan_waktu(CDate(reader("actual")))
                        If hasiltgl = 1 Then
                            notifikasi.dedline.Text = "buruan " & hasiltgl & "HARI LAGI"
                            notifikasi.dedline.BackColor = Color.Red
                            notifikasi.dedline.ForeColor = Color.White
                        ElseIf hasiltgl = 0 Then
                            notifikasi.dedline.Text = "SEKARANG TERAKHIR"
                            notifikasi.dedline.BackColor = Color.Red
                            notifikasi.dedline.ForeColor = Color.White
                        ElseIf hasiltgl < 0 Then
                            notifikasi.dedline.Text = hasiltgl & "HARI DARI TARGET <CEPETAN>"
                            notifikasi.dedline.BackColor = Color.Red
                            notifikasi.dedline.ForeColor = Color.White
                        Else
                            notifikasi.dedline.Text = hasiltgl & "HARI LAGI"
                        End If
                        notifikasi.Show()
                        notifikasi.Timer1.Stop()
                        Me.Hide()
                    End While
                End If
                If jumlah_data = 0 Then
                    notifikasi.Timer1.Stop()
                    Timer2.Enabled = True
                    If connection_stat = True Then
                        Try
                            cmd = connection.CreateCommand
                            cmd.CommandText = "update eci_app set colom_manggil = '' where keterangan = 'belum' AND destination = '" & nama_pengguna & "'"
                            cmd.ExecuteNonQuery()
                        Catch ex As Exception

                        End Try
                    End If
                Else
                    If connection_stat = True Then
                        Try
                            cmd = connection.CreateCommand
                            cmd.CommandText = "update eci_app set colom_manggil = 'tampil' where no_eci='" & notifikasi.outputnoeci.Text & "' AND destination = '" & nama_pengguna & "'"
                            cmd.ExecuteNonQuery()
                        Catch ex As Exception

                        End Try
                    End If
                End If
            End If
        End If
        Timer2.Enabled = True
    End Sub
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        am(1) = "00:00:00 AM"
        am(2) = "02:30:00 AM"
        am(3) = "04:00:00 AM"
        am(4) = "08:00:00 AM"
        am(5) = "10:00:00 AM"
        am(6) = "13:00:00 AM"
        am(7) = "16:00:00 AM"

        pm(1) = "00:00:00 PM"
        pm(2) = "02:00:00 PM"
        pm(3) = "04:00:00 PM"
        pm(4) = "08:00:00 PM"
        pm(5) = "10:00:00 PM"
        pm(6) = "13:00:00 PM"
        pm(7) = "16:00:00 PM"

        n(1) = "00:00:00"
        n(2) = "02:00:00"
        n(3) = "04:00:00"
        n(4) = "08:00:00"
        n(5) = "10:00:00"
        n(6) = "13:00:00"
        n(7) = "16:00:00"
        If ((currenttime = am(1)) Or (currenttime = am(2)) Or (currenttime = am(3)) Or (currenttime = am(4)) Or (currenttime = am(5)) Or (currenttime = am(6)) Or (currenttime = am(7)) Or (currenttime = pm(1)) Or (currenttime = pm(2)) Or (currenttime = pm(3)) Or (currenttime = pm(4)) Or (currenttime = pm(5)) Or (currenttime = pm(6)) Or (currenttime = pm(7)) Or (currenttime = n(1)) Or (currenttime = n(2)) Or (currenttime = n(3)) Or (currenttime = n(4)) Or (currenttime = n(5)) Or (currenttime = n(6)) Or (currenttime = n(7))) Then
            Timer2.Stop()
            Timer2.Enabled = False
            Call hitung()
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
        If hak = "admin" Then
            Me.Hide()
            Form1.Show()
            Form2.TextBox1.Enabled = True
            Form2.TextBox2.Enabled = True
            Form2.TextBox3.Enabled = True
            Form2.TextBox4.Enabled = True
            Form2.TextBox9.Enabled = True
            Form2.TextBox8.Enabled = True
            Form2.TextBox6.Enabled = True
            Form2.TextBox5.Enabled = False
            Form2.Button9.Enabled = True
            Form2.Button14.Enabled = True
            Form2.DateTimePicker1.Enabled = True
            Form2.DateTimePicker2.Enabled = False
            Form2.ComboBox1.Enabled = True
            Form2.ComboBox2.Enabled = True
            Form2.ComboBox3.Enabled = True
            Form2.Button1.Enabled = True
            Form2.Button2.Enabled = True
            Form2.Button3.Enabled = True
            Form2.Button4.Enabled = True
            Form2.PictureBox2.Enabled = True
        Else
            Me.Hide()
            Form2.Show()
            Form2.TextBox1.Enabled = False
            Form2.TextBox2.Enabled = False
            Form2.TextBox3.Enabled = False
            Form2.TextBox4.Enabled = False
            Form2.TextBox5.Enabled = False
            Form2.TextBox7.Enabled = False
            Form2.TextBox9.Enabled = False
            Form2.TextBox8.Enabled = False
            Form2.TextBox6.Enabled = False
            Form2.Button9.Enabled = False
            Form2.ExToolStripMenuItem.Enabled = False
            Form2.FILEToolStripMenuItem.Enabled = False
            Form2.Button14.Enabled = False
            Form2.DateTimePicker1.Enabled = False
            Form2.DateTimePicker2.Enabled = False
            Form2.ComboBox1.Enabled = False
            Form2.ComboBox2.Enabled = False
            Form2.ComboBox3.Enabled = False
            Form2.ComboBox5.Enabled = False
            Form2.Button1.Enabled = False
            Form2.Button2.Enabled = False
            Form2.Button3.Enabled = False
            Form2.Button4.Enabled = False
            Form2.PictureBox2.Enabled = False
        End If
    End Sub

    Private Sub awal_Load(sender As Object, e As EventArgs) Handles Me.Load
        koneksi()
        koneksi2()
        jumlah_data = 0
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height)
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width)
        Timer2.Start()
        Timer2.Enabled = True
        Button1.Enabled = False
        Button2.Enabled = True
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        If nama_pengguna = "" Then
            ToolStripStatusLabel1.Text = "BELUM LOGIN"
            ToolStripStatusLabel1.BackColor = Color.Red
            Button4.Enabled = False
        Else
            ToolStripStatusLabel1.Text = nama_pengguna
            ToolStripStatusLabel1.BackColor = Color.Green
            nama_pengguna = ToolStripStatusLabel1.Text
        End If
    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Sub ToolStripStatusLabel2_Click(sender As Object, e As EventArgs) Handles ToolStripStatusLabel2.Click
        login.Show()
        login.view()
    End Sub
End Class