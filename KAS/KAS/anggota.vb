Imports System.Data.OleDb
Public Class anggota
    Dim koneksistring As String
    Dim path As String = My.Application.Info.DirectoryPath + "\"
    Dim koneksidata As New OleDbConnection
    Dim perintahdata As New OleDbCommand
    Dim adapter As New OleDbDataAdapter
    Dim table As New DataTable
    Dim bacadata As OleDbDataReader
    Dim a, b As String

    Private Sub anggota_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Data Anggota"
        koneksistring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & "kas.mdb"
        Call bacadatanya()
        Call tampil()



    End Sub
    Sub tambah()
        Dim controltime As DateTime = Today
        Dim tahun As Integer
        tahun = controltime.Year
        If ComboBox2.Text = "" Then
            MsgBox("id anggota harus ada.", MsgBoxStyle.OkOnly, "info")
        Else
            Try
                koneksidata.Close()
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
                perintahdata.CommandText = "insert into anggota (id,nama) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "')"
                perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                koneksidata.Close()
                perintahdata.Dispose()
                MsgBox("succes", MsgBoxStyle.Information, "Info")
                bacadata.Close()
                koneksidata.Close()
                Call tampil()
                koneksidata.Close()
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
                perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & tahun & "')"
                perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                koneksidata.Close()
                perintahdata.Dispose()
                bacadata.Close()
                koneksidata.Close()
                Call tampilbulanan()
            Catch ex As Exception
                koneksidata.Close()
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
                perintahdata.CommandText = "update anggota set nama='" & ComboBox3.Text & "' where id='" & ComboBox2.Text & "'"
                perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                bacadata.Close()
                koneksidata.Dispose()
                MsgBox("Berhasil di update", MsgBoxStyle.DefaultButton1, "Info")
                Call tampil()
                koneksidata.Close()
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
                perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                bacadata.Close()
                koneksidata.Dispose()
                tampilbulanan()
                koneksidata.Close()
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
                perintahdata.CommandText = "update hutang set nama='" & ComboBox3.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                bacadata.Close()
                koneksidata.Dispose()
                tampilbulanan()
            End Try
            koneksidata.Close()
            Call tampil()
            Call HAPUS()
            ComboBox2.Focus()
        End If
        ComboBox2.Focus()
    End Sub
    Sub tampilbulanan()
        koneksidata.Close()
        Dim ds As DataSet = New DataSet
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        Dim perintah As String = "Select * from bulanan"
        Dim search As New OleDbDataAdapter(perintah, koneksidata)
        search.Fill(ds, "bulanan")
        DataGridView1.DataSource = ds.Tables("bulanan")
        koneksidata.Close()
    End Sub
    Sub HAPUS()
        ComboBox2.Text = ""
        ComboBox3.Text = ""
    End Sub
    Sub tampil()
        Dim ds As DataSet = New DataSet
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        Dim perintah As String = "Select * from anggota"
        Dim search As New OleDbDataAdapter(perintah, koneksidata)
        search.Fill(ds, "anggota")
        DataGridView1.DataSource = ds.Tables("anggota")
        koneksidata.Close()

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Call tambah()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Try
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "delete from anggota where id='" & ComboBox2.Text & "'"
            perintahdata.Connection = koneksidata
            Dim pertanyaan As Integer = MsgBox("Apakah anda ingin menghapus data ini? ", MsgBoxStyle.YesNo, "Pesan")
            If pertanyaan = DialogResult.Yes Then
                perintahdata.ExecuteNonQuery()
                MsgBox("Berhasil Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
                Call tampil()
                ComboBox2.Text = ""
                ComboBox3.Text = ""

            Else
                MsgBox("Gagal Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
            End If
        Catch ex As Exception
            MsgBox("Gagal menghapus data", MsgBoxStyle.Critical, "Kesalahan")
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        koneksidata.Close()
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        perintahdata.CommandText = "Select * from anggota where id='" & ComboBox2.Text & "'"
        perintahdata.Connection = koneksidata
        bacadata = perintahdata.ExecuteReader

        If bacadata.Read = True Then
            a = bacadata.Item(0).ToString
            b = bacadata.Item(1).ToString
            ComboBox3.Text = b
        Else

            bacadata.Close()

        End If
        koneksidata.Close()
    End Sub
    Sub bacadatanya()
        Try

            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "Select * from anggota"
            perintahdata.Connection = koneksidata
            bacadata = perintahdata.ExecuteReader
            If bacadata.Read = True Then

                Do
                    a = bacadata.Item(0).ToString
                    b = bacadata.Item(1).ToString
                    ComboBox2.Items.Add(a)
                    ComboBox3.Items.Add(b)
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

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        ComboBox2.Text = ""
        ComboBox3.Text = ""
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        koneksidata.Close()
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        perintahdata.CommandText = "Select * from anggota where nama='" & ComboBox3.Text & "'"
        perintahdata.Connection = koneksidata
        bacadata = perintahdata.ExecuteReader

        If bacadata.Read = True Then
            a = bacadata.Item(0).ToString
            b = bacadata.Item(1).ToString
            ComboBox2.Text = a
        Else
            bacadata.Close()

        End If
        koneksidata.Close()


    End Sub

    Private Sub DataGridView1_DoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim index As Integer
        index = e.RowIndex
        Dim pilihrow As DataGridViewRow
        pilihrow = DataGridView1.Rows(index)
        ComboBox2.Text = pilihrow.Cells(0).Value.ToString
        ComboBox3.Text = pilihrow.Cells(1).Value.ToString

    End Sub


    Private Sub ComboBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox3.KeyPress
        If e.KeyChar = Chr(13) Then
            Call tambah()
        End If
    End Sub

    Private Sub ComboBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            ComboBox3.Focus()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        koneksidata.Close()
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        perintahdata.CommandText = "Select * from anggota where id='" & ComboBox2.Text & "'"
        perintahdata.Connection = koneksidata
        bacadata = perintahdata.ExecuteReader

        If bacadata.Read = True Then
            a = bacadata.Item(0).ToString
            b = bacadata.Item(1).ToString
            ComboBox3.Text = b
        Else
            MsgBox("Data tidak ditemukan", MsgBoxStyle.Information, "info")
            bacadata.Close()

        End If
        koneksidata.Close()
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Try
            Dim pertanyaan As Integer = MsgBox("Data Anggota Akan Hilang Semua, Saya Saran kan Print Laporan Terlebih Dahulu! Hapus Data? ", MsgBoxStyle.YesNo, "Pesan")
            If pertanyaan = DialogResult.Yes Then
                If pertanyaan = DialogResult.Yes Then
                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "delete from anggota where id"
                    perintahdata.Connection = koneksidata
                    perintahdata.ExecuteNonQuery()
                    koneksidata.Close()
                    perintahdata.Dispose()
                    koneksidata.Close()
                End If
            End If
        Catch ex As Exception

        End Try
        tampil()
    End Sub

    Private Sub anggota_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Form1.Show()
    End Sub
End Class