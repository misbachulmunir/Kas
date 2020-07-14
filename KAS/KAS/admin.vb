Imports System.Data.OleDb
Public Class admin

    Dim koneksistring As String
    Dim koneksidata As New OleDbConnection
    Dim perintahdata As New OleDbCommand
    Dim adapter As New OleDbDataAdapter
    Dim table As New DataTable
    Dim bacadata As OleDbDataReader
    Dim a, b As String
    Dim path As String = My.Application.Info.DirectoryPath + "\"
    Private Sub admin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Ubah User dan Password"
        koneksistring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & "kas.mdb"
        bacadatanya()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Tidak Boleh Kosong !")
        Else

            koneksidata.Close()
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
            perintahdata.CommandText = " update [admin] set [user]='" & TextBox1.Text & "' , [pass]='" & TextBox2.Text & "'where [idkod]='" & 12 & "'"
            perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                perintahdata.Dispose()
                MsgBox("Berhasil di diatur", MsgBoxStyle.DefaultButton1, "Info")
                bacadatanya()


        End If
    End Sub

    Sub bacadatanya()
        Try
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "Select * from admin"
            perintahdata.Connection = koneksidata
            bacadata = perintahdata.ExecuteReader
            If bacadata.Read = True Then

                Do
                    a = bacadata.Item(1).ToString
                    b = bacadata.Item(2).ToString
                    TextBox1.Text = a
                    TextBox2.Text = b

                Loop Until bacadata.Read = False
                bacadata.Close()
            Else

                bacadata.Close()
            End If
            bacadata.Close()
            koneksidata.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Form1.Show()
        Me.Close()
    End Sub


End Class