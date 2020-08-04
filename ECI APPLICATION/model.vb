Imports System.Data.Odbc
Public Class modelform
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call koneksi()
        If connection_stat = True Then
            Try
                cmd = connection.CreateCommand
                cmd.CommandText = "insert into eci_app_model (model) values('" + TextBox1.Text + "')"
                cmd.ExecuteNonQuery()
                MsgBox("success :: Add model")
                Form2.ComboBox3.Text = ""
                TextBox1.Text = ""
            Catch ex As Exception
                MsgBox("GAGAL :: Add model gagal")
            End Try
        Else
            MsgBox("GAGAL :: Add model gagal")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form2.ComboBox2.Items.Clear()
        Call Form2.model()
    End Sub

    Private Sub modelform_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            cmd = connection.CreateCommand
            cmd.CommandText = "insert into eci_app_model (model) values('" + TextBox1.Text + "')"
            cmd.ExecuteNonQuery()
            MsgBox("success :: Add model")
            Form2.ComboBox3.Text = ""
            TextBox1.Text = ""
        End If
    End Sub
End Class