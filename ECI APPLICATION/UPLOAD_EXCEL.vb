Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Imports spire.xls
Public Class UPLOAD_EXCEL
    Dim connodbc As OdbcConnection
    Dim daodbc As OdbcDataAdapter
    Dim dsodbc As DataSet
    Dim cmdodbc, cmdcek, cmdodbc1, cmdcek1 As OdbcCommand
    Dim drodbc, drodbc1, cek, cek1, cekodbc, outputdestination As OdbcDataReader
    Dim connection_status As Boolean
    Dim cek_import As Boolean
    Dim a, ars, jobnumber_before, jobnumber_after As String
    Sub panggil_excel()
        Try
            Dim conna As New OleDbConnection
            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet
            Dim cmd As New OleDbCommand
            Dim dt As New DataTable

            conna = New OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;" &
                        "   data source='format-excel.xlsx';Extended Properties=Excel 12.0;")

            da = New OleDbDataAdapter("select * from [Sheet1$]", conna)
            conna.Open()
            ds.Clear()
            da.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            conna.Close()
            cek_import = True
        Catch ex As Exception
            Try
                Dim conna As New OleDbConnection
                Dim da As New OleDbDataAdapter
                Dim ds As New DataSet
                Dim cmd As New OleDbCommand
                Dim dt As New DataTable

                conna = New OleDbConnection("provider=Microsoft.ACE.OLEDB.8.0;" &
                        "   data source='format-excel.xlsx';Extended Properties=Excel 8.0;")

                da = New OleDbDataAdapter("select * from [Sheet1$]", conna)
                conna.Open()
                ds.Clear()
                da.Fill(ds)
                DataGridView1.DataSource = ds.Tables(0)
                conna.Close()
                cek_import = True
            Catch ex1 As Exception
                Try
                    Dim conna As New OleDbConnection
                    Dim da As New OleDbDataAdapter
                    Dim ds As New DataSet
                    Dim cmd As New OleDbCommand
                    Dim dt As New DataTable

                    conna = New OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;" &
                    "data source='format-excel.xlsx';Extended Properties=Excel 8.0;")

                    da = New OleDbDataAdapter("select * from [Sheet1$]", conna)
                    conna.Open()
                    ds.Clear()
                    da.Fill(ds)
                    DataGridView1.DataSource = ds.Tables(0)
                    conna.Close()
                    cek_import = True
                Catch ex2 As Exception
                    cek_import = False
                    MsgBox("Version EXCEL APP ANDA TIDAK SUPPORT")
                End Try
            End Try
        End Try
    End Sub

    Private Sub update_database_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call koneksi()
        Me.CenterToScreen()
    End Sub
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        conf()
        Dim result As Integer = MessageBox.Show("Apakah Anda Ingin Membuka Format Baru ???", "Konfirmasi File Format", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Cancel Then

        ElseIf result = DialogResult.No Then
            If System.IO.File.Exists("format-excel.xlsx") = True Then
                Try
                    Process.Start("format-excel.xlsx")
                Catch ex As Exception
                    MsgBox("File Sudah Terbuka")
                End Try
            Else
                Try
                    My.Computer.Network.DownloadFile(file_format, "format-excel.xlsx")
                    Process.Start("format-excel.xlsx")
                Catch ex As Exception
                    MsgBox("GAGAL MENDOWNLOAD FORMAT")
                End Try
            End If
        ElseIf result = DialogResult.Yes Then
            If System.IO.File.Exists("format-excel.xlsx") = True Then
                Try
                    File.Delete("format-excel.xlsx")
                    My.Computer.Network.DownloadFile(file_format, "format-excel.xlsx")

                    Process.Start("format-excel.xlsx")
                Catch ex As Exception
                    MsgBox("File Sudah Terbuka")
                End Try
            Else
                Try
                    My.Computer.Network.DownloadFile(file_format, "format-excel.xlsx")
                    Process.Start("format-excel.xlsx")
                Catch ex As Exception
                    MsgBox("GAGAL MENDOWNLOAD FORMAT")
                End Try
            End If
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim no_eci, purpose_eci, dm_no, number_part, tanggal, gambar_lama, number_part_baru, numberpart_induk, gambar_baru, part_name, line, pos, model, destination, keterangan, kolom_manggil As String
        ProgressBar1.Maximum = DGV.RowCount - 1
        ProgressBar1.Value = 0
        koneksi()
        'Try
        For hapus As Integer = 0 To DGV.RowCount - 2
            'release dat
            If DGV.Rows(hapus).Cells(0).Value Is Nothing And DGV.Rows(hapus).Cells(1).Value Is Nothing And DGV.Rows(hapus).Cells(2).Value Is Nothing And DGV.Rows(hapus).Cells(4).Value Is Nothing And DGV.Rows(hapus).Cells(5).Value Is Nothing And DGV.Rows(hapus).Cells(7).Value Is Nothing And DGV.Rows(hapus).Cells(9).Value Is Nothing And DGV.Rows(hapus).Cells(13).Value Is Nothing And DGV.Rows(hapus).Cells(14).Value Is Nothing Then
                a = a + 1
                ProgressBar1.Value += 1
                MsgBox("NO ECI = " & DGV.Rows(hapus).Cells(0).Value & " Ada Form Yang Kosong Tidak TerSimpan")
            Else
                Try
                    no_eci = DGV.Rows(hapus).Cells(0).Value
                    purpose_eci = DGV.Rows(hapus).Cells(1).Value
                    dm_no = DGV.Rows(hapus).Cells(2).Value
                    ars = DGV.Rows(hapus).Cells(3).Value
                    number_part = DGV.Rows(hapus).Cells(4).Value
                    number_part_baru = DGV.Rows(hapus).Cells(5).Value
                    jobnumber_before = DGV.Rows(hapus).Cells(6).Value
                    jobnumber_after = DGV.Rows(hapus).Cells(7).Value
                    gambar_lama = "noimage.png"
                    gambar_baru = "noimage.png"
                    part_name = DGV.Rows(hapus).Cells(10).Value
                    line = DGV.Rows(hapus).Cells(11).Value
                    pos = DGV.Rows(hapus).Cells(12).Value
                    model = DGV.Rows(hapus).Cells(13).Value
                    tanggal = Format(DGV.Rows(hapus).Cells(14).Value, "yyyy-MM-dd")
                    destination = DGV.Rows(hapus).Cells(15).Value
                    keterangan = "belum"
                    kolom_manggil = ""
                Catch ex As Exception
                End Try
                If no_eci <> "" And number_part <> "" And number_part_baru <> "" And part_name <> "" And line <> "" And model <> "" And destination <> "" Then
                    If connection_stat = True Then
                        Try
                            Dim tampil As String = "SELECT *from eci_app where planning = '" & Format(DGV.Rows(hapus).Cells(14).Value, "yyyy-MM-dd") & "' and number_part='" & number_part & "' and part_name='" & number_part_baru & "' and line='" & line & "' and pos='" & pos & "' limit 1"
                            cmdodbc = New OdbcCommand(tampil, connection)
                            drodbc = cmdodbc.ExecuteReader
                            If drodbc.Read() Then
                                a = a + 1
                                ProgressBar1.Value += 1
                            Else
                                Dim tampilin As String = "SELECT *from eci_app_ars where partno = '" & number_part & "' limit 1"
                                cmdodbc = New OdbcCommand(tampilin, connection)
                                    drodbc = cmdodbc.ExecuteReader
                                    If drodbc.Read() Then
                                        numberpart_induk = drodbc("partno_induk")
                                    End If
                                    cmd = connection.CreateCommand
                                    cmd.CommandText = "INSERT INTO `eci_app_ars`(`status`,`tanggal_entry`, `valid_part`, `partno_induk`, `partno`, `part_name`, `shop_code`, `job_no`, `qty/canban`, `rack_address`, `hole_address`, `rack_layer`, `position`, `cap_hole`, `ratio`, `part_remake`, `implementation_date`, `type`, `code_area`, `area`, `zone`, `pos`) VALUES ('sudah','" & System.DateTime.Now.ToString("yyyy/MM/dd") & "','Adaption','" & numberpart_induk & "','" & number_part_baru & "','" & part_name & "','-','" & jobnumber_after & "','-','" & ars & "','-','-','-','-','-','-','" & System.DateTime.Now.ToString("yyyy/MM/dd") & "','-','-','" & line & "','" & line & "','" & pos & "')"
                                    cmd.ExecuteNonQuery()
                                    Dim simpan As String = "insert into eci_app values('" & ars & "','" & jobnumber_before & "','" & jobnumber_after & "','" & no_eci & "','" & purpose_eci & "','" & dm_no & "','" & number_part & "','" & gambar_lama & "','" & number_part_baru & "','" & gambar_baru & "','" & part_name & "','" & line & "','" & pos & "','" & model & "','" & Format(DGV.Rows(hapus).Cells(14).Value, "yyyy-MM-dd") & "','" & destination & "','" & keterangan & "' , '" & kolom_manggil & "')"
                                    cmdodbc = New OdbcCommand(simpan, connection)
                                    cmdodbc.ExecuteNonQuery()
                                    a = a + 1
                                    ProgressBar1.Value += 1
                                End If
                        Catch ex1 As Exception
                            MsgBox(ex1.Message)
                        End Try
                    End If
                Else
                    a = a + 1
                    ProgressBar1.Value += 1
                End If
            End If
        Next
        MsgBox("data berhasil disimpan")
        DGV.Columns.Clear()
        ProgressBar1.Value = 0
        'End Try
    End Sub

    Private Sub EXITToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EXITToolStripMenuItem.Click
        Me.Hide()
        Form2.Show()
        Call Form2.bersih_db()
        Call Form2.tampildatagrid()
    End Sub
    Sub buat_wadah()

    End Sub
    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Dim value As Integer = 0
        DGV.Columns.Clear()
        DataGridView1.Columns.Clear()
        koneksi()
        panggil_excel()
        If cek_import = True Then
            Dim no_eci, purpose_eci, dm_no, number_part, tanggal, gambar_lama, number_part_baru, gambar_baru, part_name, line, pos, model, destination, keterangan, kolom_manggil As String

            With DGV
                'judul
                .Columns.Add("no_eci", "NO ECI")
                .Columns.Add("purpose_eci", "PURPOSE ECI")
                .Columns.Add("dm_number", "DM NUMBER")
                .Columns.Add("ars", "RACK ADDRESS")
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
                .Columns("No_eci").Width = 100
                    .Columns("purpose_eci").Width = 100
                .Columns("dm_number").Width = 100
                .Columns("ars").Width = 100
                .Columns("number_part").Width = 75
                .Columns("number_part_baru").Width = 75
                .Columns("job_number_lama").Width = 100
                .Columns("job_number_baru").Width = 100
                .Columns("gambar_lama").Width = 75
                    .Columns("gambar_baru").Width = 75
                    .Columns("part_name").Width = 280
                    .Columns("line").Width = 75
                    .Columns("pos").Width = 75
                    .Columns("model").Width = 75
                    .Columns("Planning").Width = 115
                .Columns("Tugas").Width = 100
                For hapus As Integer = 0 To DataGridView1.RowCount - 2
                    'release data
                    If DataGridView1.Rows(hapus).Cells(0).Value Is Nothing And DataGridView1.Rows(hapus).Cells(1).Value Is Nothing And DataGridView1.Rows(hapus).Cells(2).Value Is Nothing And DataGridView1.Rows(hapus).Cells(3).Value Is Nothing And DataGridView1.Rows(hapus).Cells(4).Value Is Nothing And DataGridView1.Rows(hapus).Cells(5).Value Is Nothing And DataGridView1.Rows(hapus).Cells(6).Value Is Nothing And DataGridView1.Rows(hapus).Cells(7).Value Is Nothing And DataGridView1.Rows(hapus).Cells(8).Value Is Nothing And DataGridView1.Rows(hapus).Cells(9).Value Is Nothing And DataGridView1.Rows(hapus).Cells(10).Value Is Nothing And DataGridView1.Rows(hapus).Cells(11).Value Is Nothing Then
                        MsgBox("NO ECI = " & DataGridView1.Rows(hapus).Cells(0).Value & " Ada Form Yang Kosong Tidak TerSimpan")
                    Else
                        Try
                            ars = ""
                            Dim tampil As String = "SELECT distinct job_no,part_name,zone,pos from eci_app_ars where partno = '" & DataGridView1.Rows(hapus).Cells(4).Value & "'  and NOT(valid_part like 'DISUSE%')"
                            cmdodbc = New OdbcCommand(tampil, connection)
                            cmdcek = New OdbcCommand(tampil, connection)
                            drodbc = cmdodbc.ExecuteReader
                            cek = cmdcek.ExecuteReader
                            If cek.Read() Then
                                While drodbc.Read()
                                    no_eci = DataGridView1.Rows(hapus).Cells(0).Value
                                    purpose_eci = DataGridView1.Rows(hapus).Cells(1).Value
                                    dm_no = DataGridView1.Rows(hapus).Cells(2).Value
                                    ars = ""
                                    Dim tampil_ars As String = "SELECT distinct rack_address from eci_app_ars where partno = '" & DataGridView1.Rows(hapus).Cells(4).Value & "'  and NOT(valid_part like 'DISUSE%') and zone='" & drodbc("zone") & "'"
                                    cmdodbc1 = New OdbcCommand(tampil_ars, connection)
                                    cmdcek1 = New OdbcCommand(tampil_ars, connection)
                                    drodbc1 = cmdodbc1.ExecuteReader
                                    cek1 = cmdcek1.ExecuteReader
                                    If cek1.Read() Then
                                        While drodbc1.Read()
                                            If ars <> "" Then
                                                ars = ars & " & " & drodbc1("rack_address")
                                            Else
                                                ars = drodbc1("rack_address")
                                            End If
                                        End While
                                    End If
                                    number_part = DataGridView1.Rows(hapus).Cells(4).Value
                                    number_part_baru = DataGridView1.Rows(hapus).Cells(5).Value
                                    jobnumber_before = drodbc("job_no")
                                    Try
                                        jobnumber_after = DataGridView1.Rows(hapus).Cells(11).Value
                                    Catch ex As Exception
                                        jobnumber_after = "-"
                                    End Try
                                    gambar_lama = "noimage.png"
                                    gambar_baru = "noimage.png"
                                    part_name = drodbc("part_name")
                                    line = drodbc("zone")

                                    pos = ""
                                    tampil_ars = "SELECT distinct pos from eci_app_ars where partno = '" & DataGridView1.Rows(hapus).Cells(4).Value & "'  and NOT(valid_part like 'DISUSE%') and zone='" & drodbc("zone") & "'"
                                    cmdodbc1 = New OdbcCommand(tampil_ars, connection)
                                    cmdcek1 = New OdbcCommand(tampil_ars, connection)
                                    drodbc1 = cmdodbc1.ExecuteReader
                                    cek1 = cmdcek1.ExecuteReader
                                    If cek1.Read() Then
                                        While drodbc1.Read()
                                            If pos <> "" Then
                                                pos = pos & " & " & drodbc1("pos")
                                            Else
                                                pos = drodbc1("pos")
                                            End If
                                        End While
                                    End If
                                    model = DataGridView1.Rows(hapus).Cells(10).Value
                                    tanggal = Format(DataGridView1.Rows(hapus).Cells(9).Value, "yyyy-MM-dd")
                                    'cek nama destiantion 
                                    Dim destination_cek As String = "SELECT *from eci_app_human where zone = '" & line & "'"
                                    cmdodbc = New OdbcCommand(destination_cek, connection)
                                    outputdestination = cmdodbc.ExecuteReader
                                    If outputdestination.Read() Then
                                        destination = outputdestination("nama")
                                    Else
                                        destination = "-"
                                    End If
                                    keterangan = "belum"
                                    kolom_manggil = ""
                                    value = value + 1
                                    .Rows.Add(no_eci, purpose_eci, dm_no, ars, number_part, number_part_baru, jobnumber_before, jobnumber_after, gambar_lama, gambar_baru, part_name, line, pos, model, DataGridView1.Rows(hapus).Cells(9).Value, destination)
                                End While
                            Else
                                value = value + 1
                                .Columns("No_eci").Width = 100
                                .Columns("purpose_eci").Width = 100
                                .Columns("dm_number").Width = 100
                                .Columns("ars").Width = 100
                                .Columns("number_part").Width = 75
                                .Columns("number_part_baru").Width = 75
                                .Columns("job_number_lama").Width = 100
                                .Columns("job_number_baru").Width = 100
                                .Columns("gambar_lama").Width = 75
                                .Columns("gambar_baru").Width = 75
                                .Columns("part_name").Width = 280
                                .Columns("line").Width = 75
                                .Columns("pos").Width = 75
                                .Columns("model").Width = 75
                                .Columns("Planning").Width = 115
                                .Columns("Tugas").Width = 100
                                .Rows.Add(DataGridView1.Rows(hapus).Cells(0).Value, DataGridView1.Rows(hapus).Cells(1).Value, DataGridView1.Rows(hapus).Cells(2).Value, "-", DataGridView1.Rows(hapus).Cells(4).Value, DataGridView1.Rows(hapus).Cells(5).Value, "-", "-", "noimage.png", "noimage.png", "-", "-", "-", "-", DataGridView1.Rows(hapus).Cells(9).Value, "-")
                            End If
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    End If
                Next
            End With
        End If
    End Sub
End Class