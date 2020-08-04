Public Class Form1
    Sub clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub
    Sub cek()
        If TextBox1.Text = "" And TextBox2.Text = "" Then
            MsgBox("silahkan isi data anda !!")
        ElseIf TextBox1.Text = "" Then
            MsgBox("Nama :: Masih Kosong")
        ElseIf TextBox2.Text = "" Then
            MsgBox("Katasandi :: Masih Kosong")
        ElseIf TextBox1.Text = "admin" And TextBox2.Text = "admineci" Then
            Form2.Show()
            Me.Hide()
            Call clear()
        Else
            MsgBox("AKUN :: SALAH.. !!!")
            Call clear()
        End If
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            Call cek()
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            Call cek()
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox2.Font = New Font("webdings", 20)
        TextBox2.PasswordChar = "="
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call cek()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        awal.Show()
    End Sub
End Class
