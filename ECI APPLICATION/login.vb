Imports System.Data.Odbc
Public Class login
    Dim reader As OdbcDataReader
    Dim adapter As OdbcDataAdapter
    Dim cmd As OdbcCommand
    Dim str As String
    'marquee
    Dim tulisan(0) As String
    Dim i, j As Integer
    Public Sub marquee()
        Timer1.Enabled = True
        tulisan(0) = "SELAMAT DATANG PROJECT ASSY 2 DI APLIKASI ECI REMIND"
        Label3.Text = tulisan(j)
    End Sub
    Sub view()
        ComboBox1.Items.Clear()
        koneksi()
        If connection_stat = True Then
            Try
                Dim tampil As String = "Select distinct(nama) FROM eci_app_human"
                cmd = New OdbcCommand(tampil, connection)
                reader = cmd.ExecuteReader
                While reader.Read()
                    ComboBox1.Items.Add(reader("nama"))
                    ComboBox1.Sorted = True
                End While
            Catch ex As Exception
                MsgBox("CHECK JARINGAN ANDA")
            End Try
        End If
    End Sub
    Sub cek()
        If ComboBox1.Text = "" Then
            MsgBox("Pilih Nama Nya dulu dong...:)")
        Else
            Call koneksi()
            If (connection_stat = True) Then
                Try
                    Dim tampil As String = "SELECT * FROM eci_app_human where nama = '" + ComboBox1.Text + "'"
                    cmd = New OdbcCommand(tampil, connection)
                    reader = cmd.ExecuteReader
                    If Not reader.HasRows Then
                        MsgBox("ANDA PENYUSUP YAH... :)")
                    Else
                        nama_pengguna = reader.Item("nama")
                        hak = reader.Item("hak_akses")
                        Me.Hide()
                        awal.Show()
                    End If
                Catch ex As Exception
                    MsgBox("Terdapat Problem di Table Eci_app_human")
                End Try
            Else
                MsgBox("CHECK JARINGAN ANDA")
            End If
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'marquee kelap klip
        If i.Equals(tulisan(j).Length) Then
            Label3.Text = ""
            If j < tulisan.Length - 1 Then
                j = j + 1
                Label3.Text = tulisan(j)
            Else
                j = 0
            End If
            i = 0
        End If
        Label3.Text = tulisan(j).Substring(0, i)
        i = i + 1
        '/marquee kelapkelip
    End Sub

    Private Sub login_Load(sender As Object, e As EventArgs) Handles Me.Load
        Timer1.Enabled = True
        view()
        Call marquee()
    End Sub
    Private Sub TextBox1_KeyDown1(sender As Object, e As KeyEventArgs)
        If (e.KeyCode = Keys.Enter) Then
            Call cek()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs)
        If (e.KeyCode = Keys.Enter) Then
            Call cek()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call cek()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub
End Class