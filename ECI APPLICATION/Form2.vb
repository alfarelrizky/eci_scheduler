Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Configuration
Imports System.Runtime.InteropServices
Imports System.Data.Odbc
Imports System.IO
Public Class Form2
    Dim reader As OdbcDataReader
    Dim adapter As OdbcDataAdapter
    Dim cmd As OdbcCommand
    Dim daodbc As OdbcDataAdapter
    Dim dsodbc As DataSet
    Dim cmdodbc, cmdodbc2, cmdcek As OdbcCommand
    Dim drodbc, drodbc2, cek, cekodbc, outputdestination As OdbcDataReader
    Dim str, samplears, samplears_line, samplears_pos, ars, jobnumber_before, jobnumber_after As String
    Dim tampil As String
    'variabel datagrid array sementara
    Dim array(300) As String
    Dim value As Integer = 0
    'file name gambar
    Dim filename_lama, filename_baru, lama As String
    Dim destination, baru, a, b As String
    Dim sample_lama As String
    Dim sample_baru As String
    'varibale datagrid array sementara

    'variabel nyimpan sementara
    Dim sample1, sample2, sample3, sample_line, sample_pos As String
    Sub tampildatagrid()
        Dim tanggal As Date = System.DateTime.Now.ToString("yyyy/MM/dd")
        Dim day As String = tanggal.Day
        Dim month As String = tanggal.Month
        Dim year As String = tanggal.Year
        koneksi()
        If connection_stat = True Then
            Try
                value = 0
                With DataGridView1
                    'judul
                    .Columns.Add("NO", "NO")
                    .Columns.Add("no_eci", "NO ECI")
                    .Columns.Add("purpose_eci", "PURPOSE ECI")
                    .Columns.Add("dm_number", "DM NUMBER")
                    .Columns.Add("ars_number", "NO ARS")
                    .Columns.Add("number_part", "NUMBER PART OLD")
                    .Columns.Add("number_part_baru", "NUMBER PART NEW")
                    .Columns.Add("job_number_lama", "JOB NUMBER OLD")
                    .Columns.Add("job_number_baru", "JOB NUMBER NEW")
                    .Columns.Add("gambar_lama", "GAMBAR PART LAMA")
                    .Columns.Add("gambar_baru", "GAMBAR PART BARU")
                    .Columns.Add("part_name", "PART NAME")
                    .Columns.Add("line", "LINE")
                    .Columns.Add("pos", "POS")
                    .Columns.Add("model", "MODEL")
                    .Columns.Add("Planning", "PLANNING")
                    .Columns.Add("Tugas", "TUGAS")

                    'ukuran lebar
                    .Columns("NO").Width = 50
                    .Columns("No_eci").Width = 100
                    .Columns("purpose_eci").Width = 100
                    .Columns("dm_number").Width = 100
                    .Columns("ars_number").Width = 100
                    .Columns("number_part").Width = 75
                    .Columns("number_part_baru").Width = 75
                    .Columns("job_number_lama").Width = 75
                    .Columns("job_number_baru").Width = 75
                    .Columns("gambar_lama").Width = 75
                    .Columns("gambar_baru").Width = 75
                    .Columns("part_name").Width = 280
                    .Columns("line").Width = 75
                    .Columns("pos").Width = 75
                    .Columns("model").Width = 75
                    .Columns("Planning").Width = 115
                    .Columns("Tugas").Width = 100

                    'list
                    Dim tampil As String = "SELECT *from eci_app where planning between '" & year & "-" & month & "-01' and '" & year & "-" & month & "-31'"
                    Using cmd As New OdbcCommand(tampil, connection)
                        Using reader As OdbcDataReader = cmd.ExecuteReader
                            While reader.Read()
                                value = value + 1
                                array(value) = reader("keterangan")
                                .Rows.Add(value, reader("no_eci"), reader("purpose_eci"), reader("dm_no"), reader("no_ars"), reader("number_part"), reader("number_part_baru"), reader("jobnumber_before"), reader("jobnumber_after"), reader("gambar_lama"), reader("gambar_baru"), reader("part_name"), reader("line"), reader("pos"), reader("model"), reader("planning"), reader("destination"))
                            End While
                        End Using
                    End Using
                End With
                DataGridView1.ReadOnly = True
            Catch ex As Exception
                MsgBox("CHECK JARINGAN ANDA")
            End Try
        Else
            MsgBox("CHECK JARINGAN ANDA")
        End If
    End Sub
    Sub tampildatagrid_terimplemen()
        koneksi()
        Dim tanggal As Date = System.DateTime.Now.ToString("yyyy/MM/dd")
        Dim day As String = tanggal.Day
        Dim month As String = tanggal.Month
        Dim year As String = tanggal.Year
        If connection_stat = True Then
            Try
                value = 0
                With DataGridView1
                    'judulload
                    .Columns.Add("NO", "NO")
                    .Columns.Add("no_eci", "NO ECI")
                    .Columns.Add("purpose_eci", "PURPOSE ECI")
                    .Columns.Add("dm_number", "DM NUMBER")
                    .Columns.Add("ars_number", "ARS NUMBER")
                    .Columns.Add("number_part", "NUMBER PART OLD")
                    .Columns.Add("number_part_baru", "NUMBER PART NEW")
                    .Columns.Add("jobnumber_lama", "JOB NUMBER OLD")
                    .Columns.Add("jobnumber_baru", "JOB NUMBER NEW")
                    .Columns.Add("gambar_lama", "GAMBAR LAMA")
                    .Columns.Add("gambar_baru", "GAMBAR BARU")
                    .Columns.Add("part_name", "PART NAME")
                    .Columns.Add("line", "LINE")
                    .Columns.Add("pos", "POS")
                    .Columns.Add("Tugas", "KONFIRMASI TUGAS")
                    .Columns.Add("Vin", "RECORD NIK")
                    .Columns.Add("model", "MODEL")
                    .Columns.Add("Suffix", "SUFFIX")
                    .Columns.Add("Liffting", "LIFFTING")
                    .Columns.Add("shift", "SHIFT")
                    .Columns.Add("planning", "PLANNING")
                    .Columns.Add("actual", "ACTUAL")
                    .Columns.Add("waktu penyelesaian", "WAKTU SELESAI")

                    'ukuran lebar
                    .Columns("NO").Width = 50
                    .Columns("No_eci").Width = 100
                    .Columns("purpose_eci").Width = 100
                    .Columns("dm_number").Width = 100
                    .Columns("ars_number").Width = 100
                    .Columns("number_part").Width = 75
                    .Columns("number_part_baru").Width = 75
                    .Columns("jobnumber_lama").Width = 75
                    .Columns("jobnumber_baru").Width = 75
                    .Columns("gambar_lama").Width = 75
                    .Columns("gambar_baru").Width = 75
                    .Columns("part_name").Width = 280
                    .Columns("line").Width = 75
                    .Columns("pos").Width = 75
                    .Columns("Tugas").Width = 100
                    .Columns("vin").Width = 75
                    .Columns("model").Width = 75
                    .Columns("Suffix").Width = 100
                    .Columns("Liffting").Width = 100
                    .Columns("Shift").Width = 100
                    .Columns("planning").Width = 70
                    .Columns("actual").Width = 70
                    .Columns("waktu penyelesaian").Width = 70

                    'list
                    Dim tampil As String = "SELECT *from eci_app_history where actual between '" & year & "-" & month & "-01' and '" & year & "-" & month & "-31'"
                    cmd = New OdbcCommand(tampil, connection)
                    reader = cmd.ExecuteReader
                    While reader.Read()
                        value = value + 1
                        array(value) = "sudah"
                        .Rows.Add(value, reader("no_eci"), reader("purpose_eci"), reader("dm_number"), reader("no_ars"), reader("number_part"), reader("number_part_baru"), reader("jobnumber_before"), reader("jobnumber_after"), reader("gambar_lama"), reader("gambar_baru"), reader("part_name"), reader("line"), reader("pos"), reader("konfir_tugas"), reader("record_nik"), reader("model"), reader("suffix_unit"), reader("liffting_unit"), reader("shift"), reader("finishing"), reader("actual"), reader("waktu penyelesaian"))
                    End While
                End With
                DataGridView1.ReadOnly = True
            Catch ex As Exception
                MsgBox("CHECK JARINGAN ANDA")
            End Try
        Else
            MsgBox("CHECK JARINGAN ANDA")
        End If
    End Sub
    Sub bersih()
        ComboBox2.Items.Clear()
        ComboBox1.Items.Clear()
        TextBox10.Text = ""
        ComboBox5.Items.Clear()
        PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
        PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
        Call model()
        ComboBox3.Items.Clear()
        Call line()
        Call user()
        pos(ComboBox3.Text)
        a = ""
        b = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        sudah.Visible = False
        belum.Visible = False
        sample1 = ""
        sample2 = ""
        sample3 = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        TextBox10.Text = ""
        ComboBox5.Text = ""
    End Sub
    Sub bersih_db()
        ComboBox2.Items.Clear()
        ComboBox1.Items.Clear()
        TextBox10.Text = ""
        ComboBox3.Items.Clear()
        ComboBox5.Items.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox6.Text = ""
        DateTimePicker1.Text = ""
        DateTimePicker2.Text = ""
        sudah.Visible = False
        belum.Visible = False
        sample1 = ""
        sample2 = ""
        sample3 = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        TextBox10.Text = ""
        ComboBox5.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
        PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
        a = ""
        b = ""
        Call model()
        Call line()
        Call user()
        pos(ComboBox3.Text)
    End Sub

    Sub user()
        ComboBox1.Items.Clear()
        Try
            Dim tampil As String = "Select distinct(nama) FROM `eci_app_human`"
            cmd = New OdbcCommand(tampil, connection)
            reader = cmd.ExecuteReader
            If Not reader.HasRows Then
            Else
                While reader.Read()
                    ComboBox1.Items.Add(reader.Item("nama"))
                    ComboBox1.Sorted = True
                End While
            End If
        Catch ex As Exception
            MsgBox("Terdapat Problem di Table Eci_app_human")
        End Try
    End Sub
    Sub model()
        ComboBox2.Items.Clear()
        Try
            Dim tampil As String = "Select distinct(model) FROM eci_app_model"
            cmd = New OdbcCommand(tampil, connection)
            reader = cmd.ExecuteReader
            If Not reader.HasRows Then
            Else
                While reader.Read()
                    ComboBox2.Items.Add(reader.Item("model"))
                    ComboBox2.Sorted = True
                End While
            End If
        Catch ex As Exception
            MsgBox("Terdapat Problem di Table Eci_app_model")
        End Try
    End Sub
    Sub line()
        ComboBox3.Items.Clear()
        koneksi()
        If connection_stat = True Then
            Try
                Dim tampil As String = "Select distinct(zone) FROM eci_app_human"
                cmd = New OdbcCommand(tampil, connection)
                reader = cmd.ExecuteReader
                While reader.Read()
                    ComboBox3.Items.Add(reader("zone"))
                    ComboBox3.Sorted = True
                End While
            Catch ex As Exception
                MsgBox("CHECK JARINGAN ANDA")
            End Try
        End If
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
        PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
        Call koneksi()
        Call koneksi2()
        'combobox
        If connection_stat = True Then
            Call user()
            Call line()
            Call pos(ComboBox2.Text)
            Call model()
            Call tampildatagrid()
            DateTimePicker1.CustomFormat = "dd-MMM-yy"
            DateTimePicker2.CustomFormat = "dd-MMM-yy"
            DateTimePicker3.CustomFormat = "yyyy-MM-dd"
            DateTimePicker4.CustomFormat = "yyyy-MM-dd"
            a = ""
            b = ""
        Else
            MsgBox("JARINGAN ANDA PROBLEM")
        End If
    End Sub
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call bersih()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim strtime As String = Format(Now, "d-M-y h-s")
        Dim file_before, file_after As String
        Dim liffting, suffix, shift As String
        Dim selisih As Integer
        koneksi()
        If connection_stat = True Then
            'Try
            If Button6.Enabled = False Then
                If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And TextBox9.Text = "" And TextBox8.Text = "" And TextBox6.Text = "" And TextBox4.Text = "" And TextBox7.Text = "" And ComboBox1.Text = "" And ComboBox2.Text = "" And ComboBox3.Text = "" And TextBox10.Text = "" And ComboBox5.Text = "" And a = "" And b = "" Then
                    MsgBox("Lengkapi data terlebih dahulu")
                    TextBox1.Focus()
                Else
                    'insert gambar ke database
                    If a = "" Or a = "noimage.png" Then
                        a = "noimage.png"
                        file_before = a
                    Else
                        file_before = "b" & strtime & "" & a
                        Try
                            If sample_lama <> "noimage.png" Then
                                File.Delete("\\10.59.5.200\eci-image\" & sample_lama)
                            End If
                            Try
                                If System.IO.File.Exists("\\10.59.5.200\eci-image\" & a) = True Then
                                    System.IO.File.Delete("\\10.59.5.200\eci-image\" & a)
                                End If
                                System.IO.File.Copy(filename_lama, "\\10.59.5.200\eci-image\" & a)
                                Try
                                    My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & a, file_before)
                                Catch ex As Exception
                                    MsgBox(ex.Message)
                                End Try
                            Catch ex As Exception
                                MsgBox("SAVE FILE GAGAL")
                            End Try
                        Catch ex As Exception
                            Try
                                If System.IO.File.Exists("\\10.59.5.200\eci-image\" & a) = True Then
                                    System.IO.File.Copy(filename_lama, "\\10.59.5.200\eci-image\" & a)
                                End If
                                System.IO.File.Copy(filename_lama, "\\10.59.5.200\eci-image\" & a)
                                Try
                                    My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & a, file_before)
                                Catch ex2 As Exception
                                    MsgBox(ex2.Message)
                                End Try
                            Catch ex1 As Exception
                                MsgBox(ex.Message)
                            End Try
                        End Try
                    End If
                    If b = "" Or b = "noimage.png" Then
                        b = "noimage.png"
                        file_after = b
                    Else
                        file_after = "a" & strtime & "" & b
                        Try
                            If sample_baru <> "noimage.png" Then
                                File.Delete("\\10.59.5.200\eci-image\" & sample_baru)
                            End If
                            System.IO.File.Exists("\\10.59.5.200\eci-image\" & b)
                            System.IO.File.Copy(filename_baru, "\\10.59.5.200\eci-image\" & b)
                            Try
                                My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & b, file_after)
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try
                        Catch ex As Exception
                            Try
                                System.IO.File.Exists("\\10.59.5.200\eci-image\" & b)
                                System.IO.File.Copy(filename_baru, "\\10.59.5.200\eci-image\" & b)
                                My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & b, file_after)
                            Catch ex1 As Exception
                                MsgBox(ex1.Message)
                            End Try
                        End Try
                    End If
                    'insert gambar kedatabase
                    With DataGridView1
                        Dim numberpart_induk As String
                        Dim tampilinc As String = "SELECT *from eci_app_ars where partno = '" & TextBox2.Text & "' and valid_part = ''"
                        cmdodbc = New OdbcCommand(tampilinc, connection)
                        cmdodbc2 = New OdbcCommand(tampilinc, connection)
                        drodbc = cmdodbc.ExecuteReader
                        drodbc2 = cmdodbc2.ExecuteReader
                        If drodbc.Read() Then
                            While drodbc2.Read
                                value = value + 1
                                numberpart_induk = drodbc("partno_induk")
                                .Rows.Insert(.NewRowIndex, value, TextBox1.Text, TextBox10.Text, TextBox7.Text, TextBox9.Text, TextBox2.Text, TextBox4.Text, TextBox8.Text, TextBox6.Text, file_before, file_after, TextBox3.Text, drodbc2("zone"), drodbc2("pos"), ComboBox2.Text, DateTimePicker1.Text, ComboBox1.Text)
                                array(value) = "belum"
                                cmd = connection.CreateCommand
                                cmd.CommandText = "insert into eci_app (no_ars,jobnumber_before,jobnumber_after,no_eci,purpose_eci,dm_no,number_part, gambar_lama, number_part_baru, gambar_baru, part_name, line ,pos, model, planning, destination, keterangan,colom_manggil) values('" + TextBox9.Text + "' ,'" + TextBox8.Text + "' ,'" + TextBox6.Text + "' ,'" + TextBox1.Text + "' , '" + TextBox10.Text + "' , '" + TextBox7.Text + "' , '" + TextBox2.Text + "' ,'" & file_before & "','" + TextBox4.Text + "','" & file_after & "','" + TextBox3.Text + "','" & drodbc2("zone") & "','" & drodbc2("pos") & "','" & ComboBox2.Text & "', '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & ComboBox1.Text & "','belum','')"
                                cmd.ExecuteNonQuery()
                                cmd = connection.CreateCommand
                                cmd.CommandText = "INSERT INTO `eci_app_ars`(`status`,`tanggal_entry`, `valid_part`, `partno_induk`, `partno`, `part_name`, `shop_code`, `job_no`, `qty/canban`, `rack_address`, `hole_address`, `rack_layer`, `position`, `cap_hole`, `ratio`, `part_remake`, `implementation_date`, `type`, `code_area`, `area`, `zone`, `pos`) VALUES ('sudah','" & System.DateTime.Now.ToString("yyyy/MM/dd") & "','Adaption','" & numberpart_induk & "','" & TextBox4.Text & "','" & TextBox3.Text & "','-','" & TextBox6.Text & "','-','" & TextBox6.Text & "','-','-','-','-','-','-','" & System.DateTime.Now.ToString("yyyy/MM/dd") & "','-','-','" & drodbc2("zone") & "','" & drodbc2("zone") & "','" & drodbc2("pos") & "')"
                                cmd.ExecuteNonQuery()
                                cmd = connection.CreateCommand
                                cmd.CommandText = "update eci_app_ars set status='sudah' , valid_part='Disuse " & System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "' where partno = '" & TextBox2.Text & "'"
                                cmd.ExecuteNonQuery()
                            End While
                        Else
                            MsgBox("PART NUMBER BEFORE TIDAK ADA DI DATA ARS / Part Before Sudah Tidak Valid (UPDATE ARS ANDA)")
                        End If
                    End With
                    TextBox1.Focus()
                    TextBox1.Clear()
                    TextBox2.Clear()
                    TextBox3.Clear()
                    TextBox4.Clear()
                    TextBox5.Clear()
                    TextBox7.Clear()
                    ComboBox1.Text = ""
                    ComboBox2.Text = ""
                    ComboBox3.Text = ""
                    TextBox10.Text = ""
                    TextBox9.Text = ""
                    TextBox8.Text = ""
                    TextBox6.Text = ""
                    ComboBox5.Text = ""
                    DateTimePicker1.Text = ""
                    DateTimePicker2.Text = ""
                End If
            ElseIf Button7.Enabled = False Then
                If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And TextBox9.Text = "" And TextBox8.Text = "" And TextBox6.Text = "" And ComboBox1.Text = "" And TextBox4.Text = "" And TextBox5.Text = "" Then
                    MsgBox("Lengkapi data terlebih dahulu")
                    TextBox1.Focus()
                Else
                    'insert gambar ke database
                    If a = "" Or a = "noimage.png" Then
                        file_before = "noimage.png"
                    Else
                        file_before = "b" & strtime & "" & a
                        Try
                            System.IO.File.Exists("\\10.59.5.200\eci-image\" & a)
                            System.IO.File.Copy(filename_lama, "\\10.59.5.200\eci-image\" & a)
                            Try
                                My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & a, file_before)
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try
                        Catch ex As Exception
                            Try
                                System.IO.File.Exists("\\10.59.5.200\eci-image\" & a)
                                System.IO.File.Copy(filename_lama, "\\10.59.5.200\eci-image\" & a)
                                Try
                                    My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & a, file_before)
                                Catch ex2 As Exception
                                    MsgBox(ex2.Message)
                                End Try
                            Catch ex1 As Exception

                            End Try
                        End Try
                    End If

                    If b = "" Or b = "noimage.png" Then
                        file_after = "noimage.png"
                    Else
                        file_after = "a" & strtime & "" & b
                        Try
                            If sample_baru <> "noimage.png" Then
                                File.Delete("\\10.59.5.200\eci-image\" & sample_baru)
                            End If
                            'File.Delete("\\10.59.5.200\eci-image\" & b)
                            System.IO.File.Exists("\\10.59.5.200\eci-image\" & b)
                            System.IO.File.Copy(filename_baru, "\\10.59.5.200\eci-image\" & b)
                            Try
                                My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & b, file_after)
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try
                        Catch ex As Exception
                            Try
                                System.IO.File.Exists("\\10.59.5.200\eci-image\" & b)
                                System.IO.File.Copy(filename_baru, "\\10.59.5.200\eci-image\" & b)
                                Try
                                    My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & b, file_after)
                                Catch ex2 As Exception
                                    MsgBox(ex2.Message)
                                End Try
                            Catch ex1 As Exception

                            End Try
                        End Try
                    End If
                    'insert gambar kedatabase
                    liffting = InputBox("LIFFTING : ", "KONFIRMASI TUGAS)")
                    suffix = InputBox("SUFFIX : ", "KONFIRMASI TUGAS)")
                    shift = InputBox("SHIFT : ", "KONFIRMASI TUGAS)")
                    selisih = DateDiff(DateInterval.Day, DateTimePicker1.Value, DateTimePicker2.Value)
                    If liffting = "" And suffix = "" And shift = "" Then
                        MsgBox("DATA BELUM LENGKAP DAN TIDAK TERSIMPAN", MessageBoxIcon.Stop)
                    Else
                        With DataGridView1
                            value = value + 1
                            .Rows.Insert(.NewRowIndex, value, TextBox1.Text, TextBox10.Text, TextBox7.Text, TextBox9.Text, TextBox2.Text, TextBox4.Text, TextBox8.Text, TextBox6.Text, file_before, file_after, TextBox3.Text, ComboBox3.Text, ComboBox5.Text, ComboBox1.Text, TextBox5.Text, ComboBox2.Text, suffix, liffting, shift, Format(DateTimePicker1.Value, "yyyy-MM-dd"), Format(DateTimePicker2.Value, "yyyy-MM-dd"), selisih)
                            array(value) = "sudah"
                        End With
                        Try
                            cmd = connection.CreateCommand
                            cmd.CommandText = "insert into eci_app_history values('" & TextBox9.Text & "','" & TextBox8.Text & "','" & TextBox6.Text & "','" + TextBox1.Text + "' , '" & TextBox10.Text & "','" & TextBox7.Text & "','" + TextBox5.Text + "', '" + TextBox2.Text + "' ,'" & file_before & "','" + TextBox4.Text + "','" & file_after & "','" + TextBox3.Text + "','" & ComboBox3.Text & "','" & ComboBox5.Text & "','" & ComboBox2.Text & "','" & ComboBox1.Text & "' , '" & liffting & "' ,'" & suffix & "' , '" & shift & "' , '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' , '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "', '" & selisih & "')"
                            cmd.ExecuteNonQuery()
                            TextBox1.Focus()
                            TextBox1.Clear()
                            TextBox2.Clear()
                            TextBox3.Clear()
                            TextBox4.Clear()
                            TextBox5.Clear()
                            TextBox9.Clear()
                            TextBox8.Clear()
                            TextBox6.Clear()
                            TextBox7.Clear()
                            ComboBox1.Text = ""
                            ComboBox2.Text = ""
                            ComboBox3.Text = ""
                            TextBox10.Text = ""
                            ComboBox5.Text = ""
                            DateTimePicker1.Text = ""
                            DateTimePicker2.Text = ""
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    End If
                End If
            End If
            a = ""
            b = ""
        Else
            MsgBox("CHECK JARINGAN ANDA")
        End If
    End Sub
    Sub pos(a)
        ComboBox5.Items.Clear()
        koneksi()
        If connection_stat = True Then
            If a = "" Then
                Try
                    Dim tampil As String = "Select distinct(pos) FROM eci_app_ars"
                    cmd = New OdbcCommand(tampil, connection)
                    reader = cmd.ExecuteReader
                    While reader.Read()
                        ComboBox5.Items.Add(reader("pos"))
                        ComboBox5.Sorted = True
                    End While
                Catch ex As Exception
                    MsgBox("CHECK JARINGAN ANDA")
                End Try
            Else
                Try
                    Dim tampil As String = "Select distinct(pos) FROM eci_app_ars where zone='" & a & "'"
                    cmd = New OdbcCommand(tampil, connection)
                    reader = cmd.ExecuteReader
                    While reader.Read()
                        ComboBox5.Items.Add(reader("pos"))
                        ComboBox5.Sorted = True
                    End While
                Catch ex As Exception
                    MsgBox("CHECK JARINGAN ANDA")
                End Try
            End If
        End If
    End Sub
    Sub pic(a)
        ComboBox3.Items.Clear()
        koneksi()
        If connection_stat = True Then
            Try
                Dim tampil As String = "Select distinct(zone) FROM eci_app_human"
                cmd = New OdbcCommand(tampil, connection)
                reader = cmd.ExecuteReader
                While reader.Read()
                    ComboBox3.Items.Add(reader("zone"))
                    ComboBox3.Sorted = True
                End While
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            If a = "" Then
                ComboBox3.SelectedItem.text = ""
            Else
                Try
                    Dim tampi1 As String = "Select distinct(zone) FROM eci_app_human  where nama = '" & a & "' "
                    cmd = New OdbcCommand(tampi1, connection)
                    reader = cmd.ExecuteReader
                    While reader.Read()
                        ComboBox3.SelectedItem.text = reader("zone")
                    End While
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Button6.Enabled = False Then
            koneksi()
            koneksi2()
            If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And ComboBox3.Text = "" And ComboBox5.Text = "" And TextBox9.Text = "" And TextBox8.Text = "" And TextBox6.Text = "" Then
                MsgBox("Pilih Data yang Ingin di Hapus .. !!")
            Else
                If DataGridView1.CurrentRow.Index <> DataGridView1.NewRowIndex Then
                    If connection_stat = True Then
                        Try
                            If sample_lama <> "noimage.png" Then
                                Try
                                    System.IO.File.Delete("\\10.59.5.200\eci-image\" & sample_lama)
                                Catch ex As Exception
                                End Try
                            End If

                            If sample_baru <> "noimage.png" Then
                                Try
                                    System.IO.File.Delete("\\10.59.5.200\eci-image\" & sample_baru)
                                Catch ex As Exception
                                End Try
                            End If
                            Try
                                DataGridView1.Rows.RemoveAt(DataGridView1.CurrentRow.Index)
                                cmd = connection.CreateCommand
                                cmd.CommandText = "delete from eci_app where no_eci = '" & TextBox1.Text & "' AND line = '" & ComboBox3.Text & "'AND number_part = '" & TextBox2.Text & "' AND part_name = '" & TextBox3.Text & "'"
                                cmd.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try

                        Catch ex As Exception
                            MsgBox("HAPUS ECI GAGAL CHECK JARINGAN ANDA")
                        End Try
                    Else
                        MsgBox("DB Not Connectino , CHECK JARINGAN ANDA")
                    End If
                End If
            End If
        ElseIf Button7.Enabled = False Then
            koneksi()
            koneksi2()
            If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And ComboBox3.Text = "" And ComboBox5.Text = "" And TextBox9.Text = "" And TextBox8.Text = "" And TextBox6.Text = "" Then
                MsgBox("Pilih Data yang Ingin di Hapus .. !!")
            Else
                If DataGridView1.CurrentRow.Index <> DataGridView1.NewRowIndex Then
                    If connection_stat = True Then
                        Try
                            If sample_lama <> "noimage.png" Then
                                If System.IO.File.Exists(sample_lama) = True Then
                                    Try
                                        System.IO.File.Delete("\\10.59.5.200\eci-image\" & sample_lama)
                                    Catch ex As Exception
                                    End Try
                                End If
                            End If

                            If sample_baru <> "noimage.png" Then
                                If System.IO.File.Exists(sample_lama) = True Then
                                    Try
                                        System.IO.File.Delete("\\10.59.5.200\eci-image\" & sample_baru)
                                    Catch ex As Exception

                                    End Try
                                End If
                            End If
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                        Try
                            DataGridView1.Rows.RemoveAt(DataGridView1.CurrentRow.Index)
                            cmd = connection.CreateCommand
                            cmd.CommandText = "delete from eci_app_history where no_eci = '" & TextBox1.Text & "' AND line = '" & ComboBox3.Text & "' AND number_part = '" & TextBox2.Text & "' and part_name = '" & TextBox3.Text & "'"
                            cmd.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    Else
                        MsgBox("DB DISCONNECT, CHECK JARINGAN ANDA")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim strtime As String = Format(Now, "d-M-y h-s")
        Dim file_before, file_after As String
        If Button6.Enabled = False Then
            If Not (a = "") Then
                If a = "noimage.png" Then
                    file_before = "noimage.png"
                Else
                    file_before = "b" & strtime & "" & a
                    Try
                        If sample_lama <> "noimage.png" Then
                            System.IO.File.Delete("\\10.59.5.200\eci-image\" & sample_lama)
                        End If
                        System.IO.File.Exists("\\10.59.5.200\eci-image\" & a)
                        System.IO.File.Copy(filename_lama, "\\10.59.5.200\eci-image\" & a)
                        Try
                            My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & a, file_before)
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    Catch ex As Exception
                        Try
                            System.IO.File.Exists("\\10.59.5.200\eci-image\" & a)
                            System.IO.File.Copy(filename_lama, "\\10.59.5.200\eci-image\" & a)
                            Try
                                My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & a, file_before)
                            Catch ex1 As Exception
                                MsgBox(ex1.Message)
                            End Try
                        Catch ex1 As Exception

                        End Try
                    End Try
                    sample_lama = file_before
                End If
            End If
            If Not (b = "") Then
                If b = "noimage.png" Then
                    file_after = "noimage.png"
                Else
                    file_after = "a" & strtime & "" & b
                    Try
                        If sample_baru <> "noimage.png" Then
                            System.IO.File.Delete("\\10.59.5.200\eci-image\" & sample_baru)
                        End If
                        System.IO.File.Exists("\\10.59.5.200\eci-image\" & b)
                        System.IO.File.Copy(filename_baru, "\\10.59.5.200\eci-image\" & b)
                        Try
                            My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & b, file_after)
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    Catch ex As Exception
                        Try
                            System.IO.File.Exists("\\10.59.5.200\eci-image\" & b)
                            System.IO.File.Copy(filename_baru, "\\10.59.5.200\eci-image\" & b)
                            Try
                                My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & b, file_after)
                            Catch ex2 As Exception
                                MsgBox(ex2.Message)
                            End Try
                        Catch ex1 As Exception

                        End Try
                    End Try
                    sample_baru = file_after
                End If
            End If
            koneksi()
            If connection_stat = True Then
                Try
                    cmd = connection.CreateCommand
                    cmd.CommandText = "update eci_app set no_ars ='" & TextBox9.Text & "',jobnumber_before ='" & TextBox8.Text & "',jobnumber_after ='" & TextBox6.Text & "',no_eci ='" & TextBox1.Text & "' ,purpose_eci='" & TextBox10.Text & "',dm_no='" & TextBox7.Text & "',part_name='" & TextBox3.Text & "',number_part = '" & TextBox2.Text & "',gambar_lama ='" & sample_lama & "',number_part_baru = '" & TextBox4.Text & "',gambar_baru ='" & sample_baru & "',line='" & ComboBox3.Text & "',pos='" & ComboBox5.Text & "',model='" & ComboBox2.Text & "',planning ='" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "',destination ='" & ComboBox1.Text & "' where no_eci ='" & sample1 & "' and number_part ='" & sample2 & "' and part_name ='" & sample3 & "' and line ='" & sample_line & "' and pos ='" & sample_pos & "' and no_ars ='" & ars & "'"
                    cmd.ExecuteNonQuery()
                    cmd = connection.CreateCommand
                    cmd.CommandText = "update eci_app_ars set rack_address='" & TextBox9.Text & "',job_no='" & TextBox8.Text & "',zone='" & ComboBox3.Text & "',pos='" & ComboBox5.Text & "',partno='" & TextBox2.Text & "', part_name='" & TextBox3.Text & "' where partno = '" & sample2 & "' and zone='" & samplears_line & "' and pos='" & samplears_pos & "' and rack_address='" & ars & "' and job_no='" & jobnumber_before & "'"
                    cmd.ExecuteNonQuery()
                    cmd = connection.CreateCommand
                    cmd.CommandText = "update eci_app_ars set rack_address='" & TextBox9.Text & "',job_no='" & TextBox6.Text & "',zone='" & ComboBox3.Text & "',pos='" & ComboBox5.Text & "',partno='" & TextBox4.Text & "', part_name='" & TextBox3.Text & "' where partno = '" & samplears & "'  and zone='" & samplears_line & "' and pos='" & samplears_pos & "'  and rack_address='" & ars & "' and job_no='" & jobnumber_after & "'"
                    cmd.ExecuteNonQuery()
                    DataGridView1.Rows.Clear()
                    DataGridView1.Columns.Clear()
                    MsgBox("Data Berhasil diUpdate")
                    bersih_db()
                    Call tampildatagrid()
                    a = ""
                    b = ""
                Catch ex As Exception
                    MsgBox("UPDATE ECI GAGAL , CHECK YOUR NETWORK")
                End Try
            Else
                MsgBox("DB DISCONNECT , CHECK JARINGAN ANDA")
            End If
        ElseIf Button7.Enabled = False Then
            If a = "noimage.png" Then
                file_before = "noimage.png"
            Else
                file_before = "b" & strtime & "" & a
                Try
                    If sample_lama <> "noimage.png" Then
                        System.IO.File.Delete("\\10.59.5.200\eci-image\" & sample_lama)
                    End If
                    System.IO.File.Exists("\\10.59.5.200\eci-image\" & a)
                    System.IO.File.Copy(filename_lama, "\\10.59.5.200\eci-image\" & a)
                    Try
                        My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & a, file_before)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                Catch ex As Exception
                    Try
                        System.IO.File.Exists("\\10.59.5.200\eci-image\" & a)
                        System.IO.File.Copy(filename_lama, "\\10.59.5.200\eci-image\" & a)
                        Try
                            My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & a, file_before)
                        Catch ex1 As Exception
                            MsgBox(ex1.Message)
                        End Try
                    Catch ex1 As Exception

                    End Try
                End Try
                sample_lama = file_before
            End If
            If Not (b = "") Then
                If b = "noimage.png" Then
                    file_after = "noimage.png"
                Else
                    file_after = "a" & strtime & "" & b
                    Try
                        If sample_baru <> "noimage.png" Then
                            System.IO.File.Delete("\\10.59.5.200\eci-image\" & sample_baru)
                        End If
                        System.IO.File.Exists("\\10.59.5.200\eci-image\" & b)
                        System.IO.File.Copy(filename_baru, "\\10.59.5.200\eci-image\" & b)
                        Try
                            My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & b, file_after)
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    Catch ex As Exception
                        Try
                            System.IO.File.Exists("\\10.59.5.200\eci-image\" & b)
                            System.IO.File.Copy(filename_baru, "\\10.59.5.200\eci-image\" & b)
                            Try
                                My.Computer.FileSystem.RenameFile("\\10.59.5.200\eci-image\" & b, file_after)
                            Catch ex2 As Exception
                                MsgBox(ex2.Message)
                            End Try
                        Catch ex1 As Exception

                        End Try
                    End Try
                    sample_baru = file_after
                End If
            End If
            koneksi()
            If connection_stat = True Then
                Try
                    cmd = connection.CreateCommand
                    cmd.CommandText = "update eci_app_history set no_ars ='" & TextBox9.Text & "',jobnumber_before ='" & TextBox8.Text & "',jobnumber_after ='" & TextBox6.Text & "',no_eci ='" & TextBox1.Text & "' ,purpose_eci='" & TextBox10.Text & "',dm_number='" & TextBox7.Text & "',part_name='" & TextBox3.Text & "',record_nik='" & TextBox5.Text & "',number_part = '" & TextBox2.Text & "',gambar_lama ='" & sample_lama & "',number_part_baru = '" & TextBox4.Text & "',gambar_baru ='" & sample_baru & "',line='" & ComboBox3.Text & "',pos='" & ComboBox5.Text & "',model='" & ComboBox2.Text & "',finishing ='" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "',actual ='" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "',konfir_tugas ='" & ComboBox1.Text & "' where no_eci ='" & sample1 & "' and number_part ='" & sample2 & "' and part_name ='" & sample3 & "' and line ='" & sample_line & "' and pos ='" & sample_pos & "' and no_ars ='" & ars & "'"
                    cmd.ExecuteNonQuery()
                    cmd = connection.CreateCommand
                    cmd.CommandText = "update eci_app_ars set rack_address='" & TextBox9.Text & "',job_no='" & TextBox8.Text & "',zone='" & ComboBox3.Text & "',pos='" & ComboBox5.Text & "',partno='" & TextBox2.Text & "', part_name='" & TextBox3.Text & "' where partno = '" & sample2 & "' and zone='" & samplears_line & "' and pos='" & samplears_pos & "' and rack_address='" & ars & "' and job_no='" & jobnumber_before & "'"
                    cmd.ExecuteNonQuery()
                    cmd = connection.CreateCommand
                    cmd.CommandText = "update eci_app_ars set rack_address='" & TextBox9.Text & "',job_no='" & TextBox6.Text & "',zone='" & ComboBox3.Text & "',pos='" & ComboBox5.Text & "',partno='" & TextBox4.Text & "', part_name='" & TextBox3.Text & "' where partno = '" & samplears & "'  and zone='" & samplears_line & "' and pos='" & samplears_pos & "'  and rack_address='" & ars & "' and job_no='" & jobnumber_after & "'"
                    cmd.ExecuteNonQuery()
                    DataGridView1.Rows.Clear()
                    DataGridView1.Columns.Clear()
                    MsgBox("Data Berhasil diUpdate")
                    bersih_db()
                    Call tampildatagrid_terimplemen()
                    a = ""
                    b = ""
                Catch ex As Exception
                    'MsgBox("UPDATE ECI GAGAL , CHECK YOUR NETWORK")
                    MsgBox(ex.Message)
                End Try
            Else
                MsgBox("DB DISCONNECT , CHECK JARINGAN ANDA")
            End If
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        pop_up_adduser.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call bersih_db()
        Call tampildatagrid()
        If hak = "admin" Then
            TextBox5.Enabled = False
            Button7.Enabled = True
            Button6.Enabled = False
            DateTimePicker2.Enabled = False
            Button2.Enabled = True
        End If
        If hak <> "admin" Then
            Button7.Enabled = True
            Button6.Enabled = False
            Button2.Enabled = False
            Button5.Enabled = False
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call bersih_db()
        Call tampildatagrid_terimplemen()
        Button7.Enabled = False
        TextBox5.Enabled = True
        Button6.Enabled = True
        DateTimePicker2.Enabled = True
        If hak <> "admin" Then
            Button2.Enabled = False
            Button6.Enabled = True
            Button5.Enabled = False
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application

        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True
            rowsTotal = DataGridView1.RowCount - 1
            colsTotal = DataGridView1.Columns.Count - 1

            With (excelWorksheet)
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView1.Columns(iC).HeaderText
                Next

                For I = 0 To rowsTotal
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView1.Rows(I).Cells(j).Value
                    Next (j)
                Next (I)

                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 10
                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With

        Catch ex As Exception
            MsgBox("Export Excel Error " & ex.Message)
        Finally
            'RELEASE ALLOACTED RESOURCES 
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing

        End Try
    End Sub

    Private Sub ExToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExToolStripMenuItem.Click
        Me.Hide()
        UPLOAD_EXCEL.Show()
    End Sub

    Private Sub EXITToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Hide()
    End Sub

    Private Sub EXITToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles EXITToolStripMenuItem.Click
        Me.Hide()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If Button6.Enabled = False Then
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            value = 0
            With DataGridView1
                'judul
                .Columns.Add("NO", "NO")
                .Columns.Add("no_eci", "NO ECI")
                .Columns.Add("purpose_eci", "PURPOSE ECI")
                .Columns.Add("dm_number", "DM NUMBER")
                .Columns.Add("ars_number", "NO ARS")
                .Columns.Add("number_part", "NUMBER PART OLD")
                .Columns.Add("number_part_baru", "NUMBER PART NEW")
                .Columns.Add("job_number_lama", "JOB NUMBER OLD")
                .Columns.Add("job_number_baru", "JOB NUMBER NEW")
                .Columns.Add("gambar_lama", "GAMBAR PART LAMA")
                .Columns.Add("gambar_baru", "GAMBAR PART BARU")
                .Columns.Add("part_name", "PART NAME")
                .Columns.Add("line", "LINE")
                .Columns.Add("pos", "POS")
                .Columns.Add("model", "MODEL")
                .Columns.Add("Planning", "PLANNING")
                .Columns.Add("Tugas", "TUGAS")

                'ukuran lebar
                .Columns("NO").Width = 50
                .Columns("No_eci").Width = 100
                .Columns("purpose_eci").Width = 100
                .Columns("dm_number").Width = 100
                .Columns("ars_number").Width = 100
                .Columns("number_part").Width = 75
                .Columns("number_part_baru").Width = 75
                .Columns("job_number_lama").Width = 75
                .Columns("job_number_baru").Width = 75
                .Columns("gambar_lama").Width = 75
                .Columns("gambar_baru").Width = 75
                .Columns("part_name").Width = 280
                .Columns("line").Width = 75
                .Columns("pos").Width = 75
                .Columns("model").Width = 75
                .Columns("Planning").Width = 115
                .Columns("Tugas").Width = 100

                'list
                'list
                koneksi()
                If connection_stat = True Then
                    Try
                        Dim tampil As String = "SELECT *from eci_app where planning >= '" & DateTimePicker3.Text & "' AND planning <='" & DateTimePicker4.Text & "'"
                        cmd = New OdbcCommand(tampil, connection)
                        reader = cmd.ExecuteReader
                        While reader.Read()
                            value = value + 1
                            array(value) = reader("keterangan")
                            .Rows.Add(value, reader("no_eci"), reader("purpose_eci"), reader("dm_no"), reader("no_ars"), reader("number_part"), reader("number_part_baru"), reader("jobnumber_before"), reader("jobnumber_after"), reader("gambar_lama"), reader("gambar_baru"), reader("part_name"), reader("line"), reader("pos"), reader("model"), reader("planning"), reader("destination"))
                        End While
                    Catch ex As Exception
                        MsgBox(ex.Message)
                        'MsgBox("GAGAL GET DATA CHECK JARINGAN ANDA")
                    End Try
                Else
                    MsgBox("GAGAL GET DATA CHECK JARINGAN ANDA")
                End If
            End With
            DataGridView1.ReadOnly = True
        ElseIf Button7.Enabled = False Then
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            value = 0
            With DataGridView1
                'judulload
                .Columns.Add("NO", "NO")
                .Columns.Add("no_eci", "NO ECI")
                .Columns.Add("purpose_eci", "PURPOSE ECI")
                .Columns.Add("dm_number", "DM NUMBER")
                .Columns.Add("ars_number", "ARS NUMBER")
                .Columns.Add("number_part", "NUMBER PART OLD")
                .Columns.Add("number_part_baru", "NUMBER PART NEW")
                .Columns.Add("jobnumber_lama", "JOB NUMBER OLD")
                .Columns.Add("jobnumber_baru", "JOB NUMBER NEW")
                .Columns.Add("gambar_lama", "GAMBAR LAMA")
                .Columns.Add("gambar_baru", "GAMBAR BARU")
                .Columns.Add("part_name", "PART NAME")
                .Columns.Add("line", "LINE")
                .Columns.Add("pos", "POS")
                .Columns.Add("Tugas", "KONFIRMASI TUGAS")
                .Columns.Add("Vin", "RECORD NIK")
                .Columns.Add("model", "MODEL")
                .Columns.Add("Suffix", "SUFFIX")
                .Columns.Add("Liffting", "LIFFTING")
                .Columns.Add("shift", "SHIFT")
                .Columns.Add("planning", "PLANNING")
                .Columns.Add("actual", "ACTUAL")
                .Columns.Add("waktu penyelesaian", "WAKTU SELESAI")

                'ukuran lebar
                .Columns("NO").Width = 50
                .Columns("No_eci").Width = 100
                .Columns("purpose_eci").Width = 100
                .Columns("dm_number").Width = 100
                .Columns("ars_number").Width = 100
                .Columns("number_part").Width = 75
                .Columns("number_part_baru").Width = 75
                .Columns("jobnumber_lama").Width = 75
                .Columns("jobnumber_baru").Width = 75
                .Columns("gambar_lama").Width = 75
                .Columns("gambar_baru").Width = 75
                .Columns("part_name").Width = 280
                .Columns("line").Width = 75
                .Columns("pos").Width = 75
                .Columns("Tugas").Width = 100
                .Columns("vin").Width = 75
                .Columns("model").Width = 75
                .Columns("Suffix").Width = 100
                .Columns("Liffting").Width = 100
                .Columns("Shift").Width = 100
                .Columns("planning").Width = 70
                .Columns("actual").Width = 70
                .Columns("waktu penyelesaian").Width = 70

                'list

                'list
                koneksi()
                If connection_stat = True Then
                    Try
                        Dim tampil As String = "SELECT *from eci_app_history  where actual >= '" & DateTimePicker3.Text & "' AND actual <='" & DateTimePicker4.Text & "'"
                        cmd = New OdbcCommand(tampil, connection)
                        reader = cmd.ExecuteReader
                        While reader.Read()
                            value = value + 1
                            array(value) = "sudah"
                            .Rows.Add(value, reader("no_eci"), reader("purpose_eci"), reader("dm_number"), reader("no_ars"), reader("number_part"), reader("number_part_baru"), reader("jobnumber_before"), reader("jobnumber_after"), reader("gambar_lama"), reader("gambar_baru"), reader("part_name"), reader("line"), reader("pos"), reader("konfir_tugas"), reader("record_nik"), reader("model"), reader("suffix_unit"), reader("liffting_unit"), reader("shift"), reader("finishing"), reader("actual"), reader("waktu penyelesaian"))
                        End While
                    Catch ex As Exception
                        MsgBox(ex.Message)
                        'MsgBox("GET DATA GAGAL , CHECK JARINGAN ANDA")
                    End Try
                Else
                    MsgBox("GET DATA GAGAL , CHECK JARINGAN ANDA")
                End If
            End With
            DataGridView1.ReadOnly = True
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        modelform.Show()
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        a = ""
        Dim result As DialogResult = OFDCopy.ShowDialog
        If (result = System.Windows.Forms.DialogResult.OK) Then
            filename_lama = OFDCopy.FileName
            Me.PictureBox3.Image = Bitmap.FromFile(filename_lama)
            a = OFDCopy.SafeFileName
        End If
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        b = ""
        Dim result As DialogResult = OFDCopy.ShowDialog
        If (result = System.Windows.Forms.DialogResult.OK) Then
            filename_baru = OFDCopy.FileName
            Me.PictureBox4.Image = Bitmap.FromFile(filename_baru)
            b = OFDCopy.SafeFileName
        End If
    End Sub

    Private Sub ComboBox3_Click(sender As Object, e As EventArgs) Handles ComboBox3.Click
        line()
    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        user()
    End Sub

    Private Sub GroupBox2_Click(sender As Object, e As EventArgs) Handles GroupBox2.Click
        model()
    End Sub

    Private Sub ComboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged
        pos(ComboBox3.Text)
    End Sub

    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click
        Dim baris = DataGridView1.CurrentRow.Index
        Call bersih()
        a = ""
        b = ""
        With DataGridView1
            Dim keterangan As Integer
            If Button6.Enabled = False Then
                Try
                    TextBox1.Text = .Item(1, baris).Value
                    sample1 = .Item(1, baris).Value
                    TextBox10.Text = .Item(2, baris).Value
                    TextBox7.Text = .Item(3, baris).Value
                    TextBox9.Text = .Item(4, baris).Value
                    ars = .Item(4, baris).Value
                    TextBox2.Text = .Item(5, baris).Value
                    sample2 = .Item(5, baris).Value
                    TextBox4.Text = .Item(6, baris).Value
                    samplears = .Item(6, baris).Value
                    TextBox8.Text = .Item(7, baris).Value
                    jobnumber_before = .Item(7, baris).Value
                    TextBox6.Text = .Item(8, baris).Value
                    jobnumber_after = .Item(8, baris).Value
                    If .Item(9, baris).Value <> "" Then
                        Try
                            PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\" & .Item(9, baris).Value)
                            sample_lama = .Item(9, baris).Value
                        Catch ex As Exception
                            PictureBox3.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
                            sample_lama = .Item(9, baris).Value
                        End Try
                    End If
                    sample_lama = .Item(9, baris).Value
                    If .Item(10, baris).Value <> "" Then
                        Try
                            PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\" & .Item(10, baris).Value)
                            sample_baru = .Item(10, baris).Value
                        Catch ex As Exception
                            PictureBox4.Image = Image.FromFile("\\10.59.5.200\eci-image\noimage.png")
                            sample_baru = "noimage.png"
                        End Try
                    End If
                    sample_baru = .Item(10, baris).Value
                    TextBox3.Text = .Item(11, baris).Value
                    sample3 = .Item(11, baris).Value
                    ComboBox3.Text = .Item(12, baris).Value
                    samplears_line = .Item(12, baris).Value
                    sample_line = .Item(12, baris).Value
                    ComboBox5.Text = .Item(13, baris).Value
                    samplears_pos = .Item(13, baris).Value
                    sample_pos = .Item(13, baris).Value
                    ComboBox2.Text = .Item(14, baris).Value
                    DateTimePicker1.Text = .Item(15, baris).Value
                    ComboBox1.Text = .Item(16, baris).Value
                    If array(keterangan) = "belum" Then
                        sudah.Visible = False
                        belum.Visible = True
                    ElseIf array(keterangan) = "sudah" Then
                        sudah.Visible = True
                        belum.Visible = False
                    End If
                Catch ex As Exception
                End Try
            Else
                Try
                    keterangan = .Item(0, baris).Value
                    TextBox1.Text = .Item(1, baris).Value
                    sample1 = .Item(1, baris).Value
                    TextBox10.Text = .Item(2, baris).Value
                    TextBox7.Text = .Item(3, baris).Value
                    TextBox9.Text = .Item(4, baris).Value
                    ars = .Item(4, baris).Value
                    TextBox2.Text = .Item(5, baris).Value
                    sample2 = .Item(5, baris).Value
                    TextBox4.Text = .Item(6, baris).Value
                    TextBox8.Text = .Item(7, baris).Value
                    TextBox6.Text = .Item(8, baris).Value
                    jobnumber_before = .Item(7, baris).Value
                    jobnumber_after = .Item(8, baris).Value
                    samplears = .Item(6, baris).Value
                    If .Item(9, baris).Value <> "" Then
                        Try
                            PictureBox3.Image = Image.FromFile("\\10.59.5.200\ECI-IMAGE\" & .Item(9, baris).Value)
                            sample_lama = .Item(9, baris).Value
                        Catch ex As Exception
                            PictureBox3.Image = Image.FromFile("\\10.59.5.200\ECI-IMAGE\noimage.png")
                            sample_lama = "noimage.png"
                        End Try
                    End If
                    If .Item(10, baris).Value <> "" Then
                        Try
                            PictureBox4.Image = Image.FromFile("\\10.59.5.200\ECI-IMAGE\" & .Item(10, baris).Value)
                            sample_baru = .Item(10, baris).Value
                        Catch ex As Exception
                            PictureBox4.Image = Image.FromFile("\\10.59.5.200\ECI-IMAGE\noimage.png")
                            sample_baru = "noimage.png"
                        End Try
                    End If
                    TextBox3.Text = .Item(11, baris).Value
                    sample3 = .Item(11, baris).Value
                    ComboBox3.Text = .Item(12, baris).Value
                    sample_line = .Item(12, baris).Value
                    ComboBox5.Text = .Item(13, baris).Value
                    sample_pos = .Item(13, baris).Value
                    ComboBox1.Text = .Item(14, baris).Value
                    TextBox5.Text = .Item(15, baris).Value
                    ComboBox2.Text = .Item(16, baris).Value
                    DateTimePicker1.Text = .Item(20, baris).Value
                    DateTimePicker2.Text = .Item(21, baris).Value
                    If array(keterangan) = "belum" Then
                        sudah.Visible = False
                        belum.Visible = True
                    ElseIf array(keterangan) = "sudah" Then
                        sudah.Visible = True
                        belum.Visible = False
                    End If
                Catch ex As Exception
                End Try
            End If
        End With
    End Sub
End Class