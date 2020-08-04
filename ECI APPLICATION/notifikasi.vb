Imports System.Data.Odbc
Public Class notifikasi
    Dim a As Integer = 0
    Dim reader As OdbcDataReader
    Dim adapter As OdbcDataAdapter
    Dim cmd As OdbcCommand
    Dim str As String
    Public lama , baru As String
    Sub bersih()
        outputnoeci.Text = ""
        outputnumberpart.Text = ""
        Button3.Text = ""
        outputpartname.Text = ""
        outputactualdate.Text = ""
    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        awal.NotifyIcon1.Visible = True
        awal.NotifyIcon1.ShowBalloonTip(3000)
        Me.Hide()
        'Call awal.agenda()
        If awal.pointer >= awal.jumlah_data Then
            awal.jumlah_data = 0
            awal.pointer = 0
            awal.Timer2.Enabled = True
            Timer1.Enabled = False
            koneksi()
            If connection_stat = True Then
                Try
                    cmd = connection.CreateCommand
                    cmd.CommandText = "update eci_app set colom_manggil = '' where keterangan = 'belum'"
                    cmd.ExecuteNonQuery()
                Catch ex As Exception

                End Try
            End If
        Else
            Timer1.Enabled = True
            koneksi()
            '            If connection_stat = True Then
            '           Try
            '          cmd = connection.CreateCommand
            '         cmd.CommandText = "update eci_app set colom_manggil = 'tampil' where no_eci='" & outputnoeci.Text & "' and number_part_baru='" & Button3.Text & "'"
            '        cmd.ExecuteNonQuery()
            '       Catch ex As Exception
            '
            'end Try
            ' End If
        End If
        Call bersih()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'a = a + 1
        'If a = 10 And awal.jumlah_data > 0 Then
        ' Timer1.Stop()
        ' Me.Show()
        ' a = 0
        ' If connection_stat = True Then
        ' cmd = connection.CreateCommand
        'cmd.CommandText = "update eci_app set colom_manggil = 'tampil' where keterangan = 'belum' AND no_eci = '" & outputnoeci.Text & "'"
        'cmd.ExecuteNonQuery()
        ' End If
        'ElseIf awal.jumlah_data <= 0 Then
        'Timer1.Stop()
        'awal.Timer2.Enabled = True
        'If connection_stat = True Then
        'cmd = connection.CreateCommand
        'cmd.CommandText = "update eci_app set colom_manggil = '' where keterangan = 'belum'"
        'cmd.ExecuteNonQuery()
        'End If
        'End If

        a = a + 1
        If a >= 10 Then
            Try
                'matiin alarm kelap kelip
                Timer2.Enabled = False
                Timer1.Enabled = False
                a = 0

                awal.pointer += 1
                outputnoeci.Text = ""
                outputnumberpart.Text = ""
                Button3.Text = ""
                Button9.Text = ""
                Button11.Text = ""
                Button14.Text = ""
                Button12.Text = ""
                Button7.Text = ""
                outputpartname.Text = ""
                outputactualdate.Text = ""
                PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
                PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")

                outputnoeci.Text = awal.no_eci(awal.pointer)
                Button11.Text = awal.line(awal.pointer)
                Button9.Text = awal.model(awal.pointer)
                outputnumberpart.Text = awal.number_part(awal.pointer)
                Button3.Text = awal.number_part_baru(awal.pointer)
                outputpartname.Text = awal.part_name(awal.pointer)
                Button14.Text = awal.rack_address(awal.pointer)
                Button12.Text = awal.jobnumber_old(awal.pointer)
                Button7.Text = awal.jobnumber_new(awal.pointer)
                outputactualdate.Text = awal.actual_date(awal.pointer)
                PictureBox3.Image = Image.FromFile(awal.gambar(awal.pointer))
                PictureBox4.Image = Image.FromFile(awal.gambar_baru(awal.pointer))
                If awal.selisih(awal.pointer) > 0 Then
                    Timer2.Enabled = True
                    dedline.Text = "H +" & awal.selisih(awal.pointer)
                    dedline.BackColor = Color.Red
                    dedline.ForeColor = Color.White
                ElseIf awal.selisih(awal.pointer) = 0 Then
                    Timer2.Enabled = True
                    dedline.Text = "SEKARANG TERAKHIR"
                    dedline.BackColor = Color.Red
                    dedline.ForeColor = Color.White
                ElseIf awal.selisih(awal.pointer) < 0 Then
                    Timer2.Enabled = True
                    dedline.Text = "H " & awal.selisih(awal.pointer)
                    dedline.BackColor = Color.Red
                    dedline.ForeColor = Color.White
                Else
                    dedline.Text = awal.selisih(awal.pointer) & " HARI"
                End If
                Me.Show()
            Catch ex As Exception

            End Try
        End If
    End Sub
    Sub remove()
        Try
            cmd = connection.CreateCommand
            cmd.CommandText = "delete from eci_app where no_eci='" & outputnoeci.Text & "' and number_part = '" & outputnumberpart.Text & "' and part_name = '" & outputpartname.Text & " ' and line = '" & Button11.Text & "'"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("DELETE ECI GAGAL")
        End Try
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim a As String
        Dim b As String
        koneksi()
        'Try
        awal.NotifyIcon1.Visible = True
            awal.NotifyIcon1.ShowBalloonTip(3000)
            Dim npk, vin, liffting, suffix, shift As String
            Dim tampil As String = "Select *from eci_app_human where nama = '" & nama_pengguna & "' limit 1"
            cmd = New OdbcCommand(tampil, connection)
            reader = cmd.ExecuteReader
            While reader.Read
                npk = reader("nama")
            End While
            a = awal.gambarlama(awal.pointer)
            b = awal.gambarbaru(awal.pointer)
            vin = InputBox("VIN : ", "KONFIRMASI TUGAS)")
            liffting = InputBox("LIFFTING : ", "KONFIRMASI TUGAS)")
            suffix = InputBox("SUFFIX : ", "KONFIRMASI TUGAS)")
            shift = InputBox("SHIFT : ", "KONFIRMASI TUGAS)")
            If vin = "" And liffting = "" And suffix = "" And shift = "" Then
                MsgBox("DATA BELUM LENGKAP DAN TIDAK TERSIMPAN", MessageBoxIcon.Stop)
                Me.Hide()
            Else
                Me.Hide()
                Dim purpose_eci, dm_number, pos, model, part_name, supplier, partnumber_induk, ptr, posptr, arsptr, dl, posdl, arsdl, trim0, postrim0, arstrim0, trim12, postrim12, arstrim12, ch1, posch1, arsch1, engine, posengine, arsengine, ch2, posch2, arsch2, final, posfinal, arsfinal, sps, possps, arssps, jundate, posjundate, arsjundate As String
                Dim tampil2 As String = "Select *from eci_app where no_eci = '" & outputnoeci.Text & "' and number_part='" & outputnumberpart.Text & "' and part_name = '" & outputpartname.Text & " ' and line = '" & Button11.Text & "'"
                cmd = New OdbcCommand(tampil2, connection)
                reader = cmd.ExecuteReader
                While reader.Read
                    purpose_eci = reader("purpose_eci")
                    dm_number = reader("dm_no")
                    pos = reader("pos")
                    model = reader("model")
            End While
            'cari id terakhir
            Dim id As Integer
            tampil2 = "Select *from eci_app_cpl order by id desc limit 1"
            cmd = New OdbcCommand(tampil2, connection)
            reader = cmd.ExecuteReader
            While reader.Read
                id = reader("id")
            End While
            tampil2 = "Select *from eci_app_cpl where partnumber ='" & outputnumberpart.Text & "'"
                cmd = New OdbcCommand(tampil2, connection)
                reader = cmd.ExecuteReader
                While reader.Read
                    partnumber_induk = reader("partnumber_induk")
                    part_name = reader("part_name")
                    supplier = reader("supplier")
                    ptr = reader("PTR")
                    posptr = reader("POS-PTR")
                    arsptr = reader("ARS-PTR")
                    dl = reader("DL")
                    posdl = reader("POS-DL")
                    arsdl = reader("ARS-DL")
                    trim0 = reader("TRIM 0")
                    postrim0 = reader("POS-TRIM 0")
                    arstrim0 = reader("ARS-TRIM 0")
                    trim12 = reader("TRIM 1.2")
                    postrim12 = reader("POS-TRIM 1.2")
                    arstrim12 = reader("ARS-TRIM 1.2")
                    ch1 = reader("CH 1")
                    posch1 = reader("POS-CH 1")
                    arsch1 = reader("ARS-CH 1")
                    engine = reader("ENGINE")
                    posengine = reader("POS-ENGINE")
                    arsengine = reader("ARS-ENGINE")
                    ch2 = reader("CH 2")
                    posch2 = reader("POS-CH 2")
                    arsch2 = reader("ARS-CH 2")
                    final = reader("FINAL")
                    posfinal = reader("POS-FINAL")
                    arsfinal = reader("ARS-FINAL")
                    sps = reader("sps")
                    possps = reader("pos-sps")
                    arssps = reader("ars-sps")
                    jundate = reader("pos-sps")
                    posjundate = reader("pos-jundate")
                    arsjundate = reader("ars-jundate")
                End While
                cmd = connection.CreateCommand
                cmd.CommandText = "delete from eci_app where number_part = '" & outputnumberpart.Text & "'"
                cmd.ExecuteNonQuery()
                cmd = connection.CreateCommand
            cmd.CommandText = "update eci_app_cpl set status='DISUSE  " & System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "' , purposeeci='E' where partnumber = '" & outputnumberpart.Text & "'"
            cmd.ExecuteNonQuery()
            id += 1
            cmd = connection.CreateCommand
            cmd.CommandText = "INSERT INTO `eci_app_cpl` VALUES ('" & id & "','','" & System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "', '" & model & "','" & dm_number & "', '" & partnumber_induk & "', '" & Button3.Text & "', '" & purpose_eci & "', '" & part_name & "', '" & supplier & "', '" & outputnoeci.Text & "', '" & Button7.Text & "', '" & ptr & "', '" & posptr & "', '" & arsptr & "', '" & dl & "', '" & posdl & "', '" & arsdl & "', '" & trim0 & "', '" & postrim0 & "' ,'" & arstrim0 & "', '" & trim12 & "', '" & postrim12 & "', '" & arstrim12 & "', '" & ch1 & "', '" & posch1 & "', '" & arsch1 & "', '" & engine & "', '" & posengine & "', '" & arsengine & "', '" & ch2 & "', '" & posch2 & "', '" & arsch2 & "', '" & final & "', '" & posfinal & "', '" & arsfinal & "', '" & sps & "', '" & possps & "', '" & arssps & "', '" & jundate & "', '" & posjundate & "', '" & arsjundate & "')"
            cmd.ExecuteNonQuery()

                Try
                    cmd = connection.CreateCommand
                    '(`no_eci`, `record_nik`, `number_part`,`gambar_lama`, `number_part_baru`,`gambar_baru`, `part_name`, `line`, `model`, `konfir_tugas`, `finishing`, `actual`, `waktu penyelesaian`)
                    cmd.CommandText = "INSERT INTO `eci_app_history` VALUES ('" & Button14.Text & "','" & Button12.Text & "','" & Button7.Text & "','" & outputnoeci.Text & "','" & purpose_eci & "','" & dm_number & "','" & vin & "', '" & outputnumberpart.Text & "','" & a & "','" & Button3.Text & "','" & b & "','" & outputpartname.Text & "','" & Button11.Text & "','" & pos & "','" & Button9.Text & "','" & npk & "','" & liffting & "', '" & suffix & "','" & shift & "','" & Format(CDate(outputactualdate.Text), "yyyy-MM-dd") & "','" & Format(Now, "yyyy-MM-dd") & "','" & awal.hasiltgl & "')"
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                Call remove()
                If awal.pointer >= awal.jumlah_data Then
                    awal.Timer2.Enabled = True
                    Timer1.Enabled = False
                    If connection_stat = True Then
                    Try
                        cmd = connection.CreateCommand
                        cmd.CommandText = "update eci_app set colom_manggil = '' where keterangan = 'belum'"
                        cmd.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                    End If
                Else
                    Timer1.Enabled = True
                    If connection_stat = True Then
                        Try
                            cmd = connection.CreateCommand
                            cmd.CommandText = "update eci_app set colom_manggil = 'tampil' where no_eci='" & outputnoeci.Text & "' and number_part_baru='" & Button3.Text & "'"
                            cmd.ExecuteNonQuery()
                        Catch ex As Exception

                        End Try
                    End If
                End If
            End If
            awal.konfirmasi = "ok"
            'Catch ex As Exception
        'MsgBox("CHECK JARINGAN ANDA , TIDAK BISA MENGKONFIRMASI")
        'End Try
    End Sub
    Private Sub notifikasi_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height)
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width)
        Timer2.Enabled = True
        Timer2.Start()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub outputnoeci_Click(sender As Object, e As EventArgs) Handles outputnoeci.Click

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    End Sub

    Private Sub outputactualdate_Click(sender As Object, e As EventArgs) Handles outputactualdate.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub

    Private Sub outputpartname_Click(sender As Object, e As EventArgs) Handles outputpartname.Click

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If awal.hasiltgl <= 1 Then
            If dedline.BackColor = Color.Red Then
                dedline.BackColor = Color.Green
                dedline.ForeColor = Color.Black
            ElseIf dedline.BackColor = Color.Green Then
                dedline.BackColor = Color.Red
                dedline.ForeColor = Color.White
            End If
        End If
    End Sub
End Class