Public Class Form3
    Dim hasil_sementara As Integer
    Dim angka1 As Integer
    Private Sub tambah_Click(sender As Object, e As EventArgs) Handles tambah.Click
        angka1 = CInt(angka.Text)
        Label1.Text = Label1.Text + angka.Text + "+"
        angka.Text = ""
    End Sub

    Private Sub samadengan_Click(sender As Object, e As EventArgs) Handles samadengan.Click
        Label1.Text = CStr(angka1) + "+" + angka.Text
        hasil.Text = angka1 + CInt(angka.Text)
        angka.Text = ""
    End Sub

    Private Sub hasil_TextChanged(sender As Object, e As EventArgs) Handles hasil.TextChanged

    End Sub
End Class