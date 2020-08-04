Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Data.Odbc
Module Module1
    Public cmd As OdbcCommand
    Public rd As OdbcDataReader
    Public rd2 As OdbcDataReader
    Public adapter As OdbcDataAdapter
    Public dataset As New DataSet
    Public connection As OdbcConnection
    Public connection2 As OdbcConnection
    Public reader As OdbcDataReader
    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String,
            ByVal lpKeyName As String,
            ByVal lpDefault As String,
            ByVal lpReturnedString As StringBuilder,
            ByVal nSize As Integer,
            ByVal lpFileName As String) As Integer
    Dim objIniFile As New clsIni(My.Application.Info.DirectoryPath & "\conf.txt")
    Public server, uid, pwd, databases, port, nama, server2, uid2, pwd2, databases2, port2, file_format As String
    Public connection_stat, connection_stat2, connection_content As Boolean
    Public nama_pengguna, hak As String
    Public Sub conf()
        'koneksi
        server = objIniFile.GetString("conn", "server", "")
        uid = objIniFile.GetString("conn", "uid", "")
        pwd = objIniFile.GetString("conn", "pwd", "")
        databases = objIniFile.GetString("conn", "databases", "")
        port = objIniFile.GetString("conn", "port", "")

        server2 = objIniFile.GetString("conn_pis2", "server2", "")
        uid2 = objIniFile.GetString("conn_pis2", "uid2", "")
        pwd2 = objIniFile.GetString("conn_pis2", "pwd2", "")
        databases2 = objIniFile.GetString("conn_pis2", "databases2", "")
        port2 = objIniFile.GetString("conn_pis2", "port2", "")

        File_format = objIniFile.GetString("file_format", "file", "")
        nama_pengguna = objIniFile.GetString("user", "nama", "")
    End Sub
    Sub koneksi()
        Call conf()
        Try
            connection = New OdbcConnection("Driver=MySQL ODBC 5.1 Driver;SERVER=" & server & ";UID=" & uid & ";pwd=" & pwd & ";DATABASE=" & databases & ";PORT=" & port & "")
            'membuka koneksi
            If connection.State = ConnectionState.Closed Then
                connection.Open()
            End If
            connection_stat = True
        Catch ex As Exception
            connection_stat = False
        End Try
        If nama_pengguna = "" Then
            MsgBox("Anda Belum Login")
        Else
            If (connection_stat = True) Then
                Try
                    Dim tampil As String = "SELECT * FROM eci_app_human where nama = '" + nama_pengguna + "'"
                    cmd = New OdbcCommand(tampil, connection)
                    reader = cmd.ExecuteReader
                    If Not reader.HasRows Then
                        MsgBox("Nama Pengguna PC ECI Tidak Valid :)")
                    Else
                        nama_pengguna = reader.Item("nama")
                        hak = reader.Item("hak_akses")
                    End If
                Catch ex As Exception
                    MsgBox("Terdapat Problem di Table Eci_app_human")
                End Try
            Else
                MsgBox("CHECK JARINGAN ANDA")
            End If
        End If
    End Sub
    Sub koneksi2()
        Call conf()
        Try
            connection2 = New OdbcConnection("Driver=MySQL ODBC 5.1 Driver;SERVER=" & server2 & ";UID=" & uid2 & ";pwd=" & pwd2 & ";DATABASE=" & databases2 & ";PORT=" & port2 & "")
            'membuka koneksi
            If connection2.State = ConnectionState.Closed Then
                connection2.Open()
            End If
            connection_stat2 = True
        Catch ex As Exception
            connection_stat2 = False
        End Try
    End Sub
End Module