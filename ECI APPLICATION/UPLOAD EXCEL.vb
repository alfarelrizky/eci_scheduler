Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Data.OleDb
Imports System.Data.Odbc
Public Class UPLOAD_EXCEL
    Dim connodbc As OdbcConnection
    Dim daodbc As OdbcDataAdapter
    Dim dsodbc As DataSet
    Dim cmdodbc As OdbcCommand
    Dim drodbc As OdbcDataReader
    Dim connection_status As Boolean
    Dim a As String
    Sub Koneksimysql()
        Try
            connodbc = New OdbcConnection("Driver=MySQL ODBC 5.1 Driver;SERVER=10.59.4.107; UID=farel;pwd=Cryptonesia22;DATABASE=picking_lamp_trial_backup;PORT=3306")
            'connodbc = New OdbcConnection("Driver=MySQL ODBC 5.1 Driver;SERVER=localhost; UID=root;pwd=;DATABASE=picking_lamp_trial;PORT=3306")
            If connodbc.State = ConnectionState.Closed Then
                connodbc.Open()
                ToolStripStatusLabel1.Text = "Success Connection ..."
                connection_status = True
            End If
        Catch ex As Exception
            ToolStripStatusLabel1.Text = "Connection Error..."
            connection_status = False
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim conn As OleDbConnection
        Dim da As OleDbDataAdapter
        Dim ds As New DataSet
        Dim cmd As OleDbCommand
        Dim dt As New DataTable

        On Error Resume Next
        OpenFileDialog1.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()

        conn = New OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;" &
                    "data source='" & OpenFileDialog1.FileName & "';Extended Properties=Excel 8.0;")

        da = New OleDbDataAdapter("select * from [Sheet1$]", conn)
        conn.Open()
        ds.Clear()
        da.Fill(ds)
        DGV.DataSource = ds.Tables(0)
        conn.Close()
    End Sub

    Private Sub update_database_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Koneksimysql()
        Me.CenterToScreen()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
End Class