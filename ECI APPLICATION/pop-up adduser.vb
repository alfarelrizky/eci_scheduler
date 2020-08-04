Imports System.Data.Odbc
Public Class pop_up_adduser
    Dim str As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call koneksi()
        If connection_stat = True Then
            Try
                cmd = connection.CreateCommand
                cmd.CommandText = "insert into eci_app_human (nama , hak_akses) values('" + TextBox1.Text + "' , '" + ComboBox1.Text + "')"
                cmd.ExecuteNonQuery()
                MsgBox("success :: Add user")
                TextBox1.Text = ""
                ComboBox1.Text = ""
            Catch ex As Exception
                MsgBox("GAGAL :: Add user gagal")
            End Try
        Else
            MsgBox("GAGAL :: Add user gagal")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form2.ComboBox2.Items.Clear()
        Call Form2.model()
    End Sub
    Private Sub pop_up_adduser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
    End Sub
End Class