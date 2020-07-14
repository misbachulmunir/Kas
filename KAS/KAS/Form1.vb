
Imports System.Data.OleDb
Public Class Form1

    Dim koneksistring As String
    Dim path As String = My.Application.Info.DirectoryPath + "\"
    Dim koneksidata As New OleDbConnection
    Dim perintahdata As New OleDbCommand
    Dim adapter As New OleDbDataAdapter
    Dim table As New DataTable
    Dim bacadata As OleDbDataReader
    Dim hutang, pengeluaranlain, kasbulanan, pemasukanlain, cmb1, cmb2, cmb3, cmb4, cmb5, cmb6 As Boolean
    Dim a, b, pndh1, pndh2, pndh3, pndh5, pndh6, an As String
    Dim pndh4 As Double
    Dim nominal As Double

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "MENU UTAMA"
        koneksistring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & "kas.mdb"
        CheckBox1.Enabled = False
        tampilbulanan()
        tampilhutang()
        tampilpemasukanlain()
        tampilpengeluaranlain()

    End Sub
    Sub bacadataanggota()
        Try
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "Select * from anggota"
            perintahdata.Connection = koneksidata
            bacadata = perintahdata.ExecuteReader
            If bacadata.Read = True Then

                Do
                    a = bacadata.Item(0).ToString
                    b = bacadata.Item(1).ToString
                    ComboBox7.Items.Add(b)
                    ComboBox12.Items.Add(a)
                    ComboBox6.Items.Add(a)
                    ComboBox9.Items.Add(a)
                    ComboBox8.Items.Add(b)
                    ComboBox10.Items.Add(b)
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
    Sub tampilhutang()
        koneksidata.Close()
        Dim ds As DataSet = New DataSet
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        Dim perintah As String = "Select * from hutang"
        Dim search As New OleDbDataAdapter(perintah, koneksidata)
        search.Fill(ds, "hutang")
        DataGridView4.DataSource = ds.Tables("hutang")
        koneksidata.Close()
    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        ComboBox6.Items.Clear()
        ComboBox8.Items.Clear()
        bacadataanggota()
        tampilpengeluaranlain()
        pengeluaranlain = True
        hutang = False
        pemasukanlain = False
        tampilpengeluaranlain()
        kasbulanan = False
    End Sub
    Sub simpanhutang()


        If ComboBox12.Text = "" Then
            MsgBox("id anggota harus ada.", MsgBoxStyle.OkOnly, "info")
        Else
            Try
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
                perintahdata.CommandText = "insert into hutang (id_anggota,nama,tanggal,nominal) values ('" & ComboBox12.Text & "','" & ComboBox7.Text & "','" & DateTimePicker3.Text & "','" & TextBox6.Text & "')"
                perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                koneksidata.Close()
                perintahdata.Dispose()
                MsgBox("succes", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                Call tampilhutang()
            Catch ex As Exception
                If TextBox6.Text = "" Then
                    MsgBox("Nominal harus di isi atau tidak boleh nol ")
                ElseIf TextBox6.Text > 0 Or TextBox6.Text < 0 Then

                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "select * from hutang where id_anggota='" & ComboBox12.Text & "'"
                    perintahdata.Connection = koneksidata
                    bacadata = perintahdata.ExecuteReader
                    If bacadata.Read = True Then
                        nominal = bacadata.Item(3)
                        bacadata.Close()
                    Else

                        bacadata.Close()
                    End If
                    bacadata.Close()
                    koneksidata.Close()
                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "update hutang set nama='" & ComboBox7.Text & "', tanggal='" & DateTimePicker3.Text & "',nominal='" & TextBox6.Text + nominal & "' where id_anggota='" & ComboBox12.Text & "'"
                    perintahdata.Connection = koneksidata
                    perintahdata.ExecuteNonQuery()
                    MsgBox("Berhasil di update", MsgBoxStyle.DefaultButton1, "Info")
                Else
                    MsgBox("Tidak boleh nol")
                End If
                ambilnominalhutang()
                If pndh4 = 0 Then
                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "delete from hutang where id_anggota='" & ComboBox12.Text & "'"
                    perintahdata.Connection = koneksidata
                    Dim pertanyaan As Integer = MsgBox("Sodara dengan ID " + ComboBox12.Text + " bernama " + ComboBox7.Text + " Telah Lunas data akan dihapus", MsgBoxStyle.YesNo, "info")
                    If pertanyaan = DialogResult.Yes Then
                        perintahdata.ExecuteNonQuery()
                        koneksidata.Close()
                        perintahdata.Dispose()
                        Call tampilhutang()

                    Else
                        MsgBox("Gagal Menghapus", MsgBoxStyle.Information, "Info")
                        koneksidata.Close()
                        perintahdata.Dispose()
                    End If
                End If

                Call tampilhutang()
            End Try
            koneksidata.Close()
            Call tampilhutang()

            ComboBox12.Focus()
        End If
        koneksidata.Close()
        TextBox6.Enabled = True
    End Sub
    Sub simpanpengeluaranlain()

        If TextBox1.Text = "" Then
            MsgBox("id pengeluaran harus ada.", MsgBoxStyle.OkOnly, "info")
        Else
            Try
                koneksidata.Close()
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
                perintahdata.CommandText = "insert into pengeluaranlain (id_pengeluaranlain,id_anggota,nama,nama_pengeluaran,tanggal,nominal) values ('" & TextBox1.Text & "','" & ComboBox6.Text & "','" & ComboBox8.Text & "','" & TextBox5.Text & "','" & DateTimePicker4.Text & "','" & TextBox7.Text & "')"
                perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                koneksidata.Close()
                perintahdata.Dispose()
                MsgBox("succes", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                Call tampilpengeluaranlain()
            Catch ex As Exception

                If TextBox7.Text = "" Then
                    MsgBox("Nominal harus di isi atau tidak boleh nol ")
                ElseIf TextBox7.Text > 0 Or TextBox7.Text < 0 Then
                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "update pengeluaranlain set id_anggota='" & ComboBox6.Text & "',nama='" & ComboBox8.Text & "', nama_pengeluaran='" & TextBox5.Text & "', tanggal='" & DateTimePicker4.Text & "',nominal='" & TextBox7.Text & "' where id_pengeluaranlain='" & TextBox1.Text & "'"
                    perintahdata.Connection = koneksidata
                    perintahdata.ExecuteNonQuery()
                    MsgBox("Berhasil di update", MsgBoxStyle.DefaultButton1, "Info")
                Else
                    MsgBox("Tidak boleh nol")
                End If

            End Try
        End If
        Call tampilpengeluaranlain()
        koneksidata.Close()
        TextBox1.Focus()
        batalpengeluaranlain()
    End Sub
    Sub simpanpemasukanlain()
        If TextBox2.Text = "" Then
            MsgBox("id pemasukan harus ada.", MsgBoxStyle.OkOnly, "info")
        Else
            Try
                koneksidata.Close()
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
                perintahdata.CommandText = "insert into pemasukanlain (id_pemasukanlain,id_anggota,nama,nama_pemasukan,tanggal,nominal) values ('" & TextBox2.Text & "','" & ComboBox9.Text & "','" & ComboBox10.Text & "','" & TextBox4.Text & "','" & DateTimePicker2.Text & "','" & TextBox3.Text & "')"
                perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                koneksidata.Close()
                perintahdata.Dispose()
                MsgBox("succes", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                Call tampilpengeluaranlain()
            Catch ex As Exception

                If TextBox3.Text = "" Then
                    MsgBox("Nominal harus di isi atau tidak boleh nol ")
                ElseIf TextBox3.Text > 0 Or TextBox3.Text < 0 Then
                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "update pemasukanlain set id_anggota='" & ComboBox9.Text & "',nama='" & ComboBox10.Text & "', nama_pemasukan='" & TextBox4.Text & "', tanggal='" & DateTimePicker2.Text & "',nominal='" & TextBox3.Text & "' where id_pemasukanlain='" & TextBox2.Text & "'"
                    perintahdata.Connection = koneksidata
                    perintahdata.ExecuteNonQuery()
                    MsgBox("Berhasil di update", MsgBoxStyle.DefaultButton1, "Info")
                Else
                    MsgBox("Tidak boleh nol")
                End If

            End Try
        End If
        Call tampilpemasukanlain()
        koneksidata.Close()
        TextBox2.Focus()
        batalpemasukanlain()
    End Sub

    Sub batalpengeluaranlain()
        TextBox1.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        ComboBox8.Text = ""
        ComboBox6.Text = ""
        TextBox1.Enabled = True
        TextBox5.Enabled = True
        DateTimePicker4.Enabled = True
        TextBox7.Enabled = True
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If hutang = True Then
            Call simpanhutang()
            cancelhutang()
            Label25.Visible = False
            CheckBox1.Checked = False
            saldong()


        ElseIf pengeluaranlain = True
            simpanpengeluaranlain()
            saldong()

        ElseIf pemasukanlain = True
            simpanpemasukanlain()
            batalpemasukanlain()
            Button3.Enabled = True
            saldong()
        ElseIf kasbulanan = True
            If ComboBox5.Text = "" Then
                MsgBox("Nominal Tidak Boleh Kosong !")
            Else
                simpanbulanan()
                batalbulanan()
                saldong()
            End If

        End If
    End Sub
    Sub simpanbulanan()
        If ComboBox2.Text = "" And ComboBox2.Enabled = True Then
            MsgBox("id kosong !")
        ElseIf CheckBox2.CheckState = CheckState.Checked Then
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            If ComboBox4.Text = "Januari" Then
                perintahdata.CommandText = " update bulanan set tahun='" & ComboBox1.Text & "', januari='" & ComboBox5.Text & "' "
            ElseIf ComboBox4.Text = "Februari" Then
                perintahdata.CommandText = "update bulanan set tahun='" & ComboBox1.Text & "', februari='" & ComboBox5.Text & "'"
            ElseIf ComboBox4.Text = "Maret" Then
                perintahdata.CommandText = "update bulanan set tahun='" & ComboBox1.Text & "', maret='" & ComboBox5.Text & "' "
            ElseIf ComboBox4.Text = "April" Then
                perintahdata.CommandText = "update bulanan set  tahun='" & ComboBox1.Text & "', april='" & ComboBox5.Text & "'"
            ElseIf ComboBox4.Text = "Mei" Then
                perintahdata.CommandText = "update bulanan set tahun='" & ComboBox1.Text & "', mei='" & ComboBox5.Text & "' "
            ElseIf ComboBox4.Text = "Juni" Then
                perintahdata.CommandText = "update bulanan set  tahun='" & ComboBox1.Text & "', juni='" & ComboBox5.Text & "'"
            ElseIf ComboBox4.Text = "Juli" Then
                perintahdata.CommandText = "update bulanan set tahun='" & ComboBox1.Text & "', juli='" & ComboBox5.Text & "' "
            ElseIf ComboBox4.Text = "Agustus" Then
                perintahdata.CommandText = "update bulanan set tahun='" & ComboBox1.Text & "', agustus='" & ComboBox5.Text & "' "
            ElseIf ComboBox4.Text = "September" Then
                perintahdata.CommandText = "update bulanan set tahun='" & ComboBox1.Text & "', september='" & ComboBox5.Text & "'"
            ElseIf ComboBox4.Text = "Oktober" Then
                perintahdata.CommandText = "update bulanan set tahun='" & ComboBox1.Text & "', oktober='" & ComboBox5.Text & "'"
            ElseIf ComboBox4.Text = "November" Then
                perintahdata.CommandText = "update bulanan set tahun='" & ComboBox1.Text & "', november='" & ComboBox5.Text & "'"
            ElseIf ComboBox4.Text = "Desember" Then
                perintahdata.CommandText = "update bulanan set  tahun='" & ComboBox1.Text & "', desember='" & ComboBox5.Text & "'"
            End If
            perintahdata.Connection = koneksidata
            perintahdata.ExecuteNonQuery()
            koneksidata.Dispose()
            tampilbulanan()
        Else
            Try
                koneksidata.Close()
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
                If ComboBox4.Text = "Januari" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,januari) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "Februari" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,februari) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "Maret" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,maret) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "April" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,april) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "Mei" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama],tahun,mei) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "Juni" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,juni] values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "Juli" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,juli) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "Agustus" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,agustus) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "September" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,september) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "Oktober" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,oktober) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "November" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,november) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                ElseIf ComboBox4.Text = "Desember" Then
                    perintahdata.CommandText = "insert into bulanan (id_anggota,nama,tahun,desember) values ('" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox5.Text & "')"
                End If


                perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                koneksidata.Close()
                perintahdata.Dispose()
                MsgBox("succes", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                Call tampilpengeluaranlain()
                koneksidata.Close()
                ComboBox2.Focus()
                batalbulanan()
            Catch ex As Exception
                koneksidata.Close()
                koneksidata.ConnectionString = koneksistring
                koneksidata.Open()
                If ComboBox4.Text = "Januari" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', januari='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "Februari" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', februari='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "Maret" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', maret='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "April" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', april='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "Mei" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', mei='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "Juni" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', juni='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "Juli" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', juli='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "Agustus" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', agustus='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "September" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "',tahun='" & ComboBox1.Text & "', september='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "Oktober" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', oktober='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "November" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', november='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                ElseIf ComboBox4.Text = "Desember" Then
                    perintahdata.CommandText = "update bulanan set nama='" & ComboBox3.Text & "', tahun='" & ComboBox1.Text & "', desember='" & ComboBox5.Text & "' where id_anggota='" & ComboBox2.Text & "'"
                End If
                perintahdata.Connection = koneksidata
                perintahdata.ExecuteNonQuery()
                MsgBox("Berhasil di update", MsgBoxStyle.DefaultButton1, "Info")

            End Try
            Call tampilbulanan()
            koneksidata.Close()
            ComboBox2.Focus()
            batalbulanan()
        End If
    End Sub
    Sub batalbulanan()
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        CheckBox2.CheckState = CheckState.Unchecked
    End Sub

    Private Sub PrintLaporanToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PrintLaporanToolStripMenuItem1.Click
        anggota.Show()
        Me.Hide()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If hutang = True Then
            If ComboBox12.Text = "" Then
                MsgBox("id kosong")
            Else
                hapushutang()
                cancelhutang()
                Label25.Visible = False
                CheckBox1.Checked = False
            End If
            saldong()
        ElseIf pengeluaranlain = True
            If TextBox1.Text = "" Then
                MsgBox("id kosong")
            Else
                hapuspengeluaranlain()
            End If
            saldong()
        ElseIf pemasukanlain = True
            If TextBox2.Text = "" Then
                MsgBox("id kosong")
            Else
                hapuspemasuklain()
                Button3.Enabled = True
                batalpemasukanlain()
            End If
            saldong()
        ElseIf kasbulanan = True
            If ComboBox2.Text = "" Then
                MsgBox("id kosong")

            Else
                hapusbulanan()
                batalbulanan()
            End If
            saldong()
        End If
    End Sub
    Sub hapusbulanan()
        Try
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "delete from bulanan where id_anggota='" & ComboBox2.Text & "'"
            perintahdata.Connection = koneksidata
            Dim pertanyaan As Integer = MsgBox("Apakah anda ingin menghapus data ini? ", MsgBoxStyle.YesNo, "Pesan")
            If pertanyaan = DialogResult.Yes Then
                perintahdata.ExecuteNonQuery()
                MsgBox("Berhasil Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
                Call tampilpengeluaranlain()

            Else
                MsgBox("Gagal Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
            End If
        Catch ex As Exception
            MsgBox("Gagal menghapus data", MsgBoxStyle.Critical, "Kesalahan")
        End Try
        tampilbulanan()
        batalbulanan()
    End Sub
    Sub hapuspengeluaranlain()
        Try
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "delete from pengeluaranlain where id_pengeluaranlain='" & TextBox1.Text & "'"
            perintahdata.Connection = koneksidata
            Dim pertanyaan As Integer = MsgBox("Apakah anda ingin menghapus data ini? ", MsgBoxStyle.YesNo, "Pesan")
            If pertanyaan = DialogResult.Yes Then
                perintahdata.ExecuteNonQuery()
                MsgBox("Berhasil Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
                Call tampilpengeluaranlain()

            Else
                MsgBox("Gagal Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
            End If
        Catch ex As Exception
            MsgBox("Gagal menghapus data", MsgBoxStyle.Critical, "Kesalahan")
        End Try
        TextBox7.Enabled = True
        batalpengeluaranlain()
    End Sub
    Sub hapushutang()
        Try
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "delete from hutang where id_anggota='" & ComboBox12.Text & "'"
            perintahdata.Connection = koneksidata
            Dim pertanyaan As Integer = MsgBox("Apakah anda ingin menghapus data ini? ", MsgBoxStyle.YesNo, "Pesan")
            If pertanyaan = DialogResult.Yes Then
                perintahdata.ExecuteNonQuery()
                MsgBox("Berhasil Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
                Call tampilhutang()

            Else
                MsgBox("Gagal Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
            End If
        Catch ex As Exception
            MsgBox("Gagal menghapus data", MsgBoxStyle.Critical, "Kesalahan")
        End Try
        TextBox6.Enabled = True
    End Sub

    Private Sub DataGridView4_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellDoubleClick
        Dim index As Integer
        index = e.RowIndex
        Dim pilihrow As DataGridViewRow
        pilihrow = DataGridView4.Rows(index)

        ComboBox12.Text = pilihrow.Cells(0).Value.ToString
        ComboBox7.Text = pilihrow.Cells(1).Value.ToString
        DateTimePicker3.Text = pilihrow.Cells(2).Value.ToString
        TextBox6.Text = pilihrow.Cells(3).Value.ToString

        Button5.Enabled = False

        ComboBox12.Enabled = False
        ComboBox7.Enabled = False
        DateTimePicker3.Enabled = False
        ambilnominalhutang()
        CheckBox1.Enabled = True

    End Sub
    Sub ambilnominalhutang()
        koneksidata.Close()
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        perintahdata.CommandText = "Select * from hutang where id_anggota ='" & ComboBox12.Text & "'"
        perintahdata.Connection = koneksidata
        bacadata = perintahdata.ExecuteReader

        If bacadata.Read = True Then
            pndh4 = bacadata.Item(3).ToString
            Label25.Text = pndh4
        End If
    End Sub
    Private Sub TabPage7_Enter(sender As Object, e As EventArgs) Handles TabPage7.Enter
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()
        bacadataanggota()
        tampilbulanan()
        kasbulanan = True
        pemasukanlain = False
        pengeluaranlain = False
        Dim controltime As DateTime = Today
        Dim tahun, bulan As Integer
        tahun = controltime.Year
        bulan = controltime.Month
        ComboBox1.Text = tahun
        If bulan = 1 Then
            ComboBox4.Text = "Januari"
        ElseIf bulan = 2 Then
            ComboBox4.Text = "Februari"
        ElseIf bulan = 3 Then
            ComboBox4.Text = "Maret"
        ElseIf bulan = 4 Then
            ComboBox4.Text = "April"
        ElseIf bulan = 5 Then
            ComboBox4.Text = "Mei"
        ElseIf bulan = 6 Then
            ComboBox4.Text = "Juni"
        ElseIf bulan = 7 Then
            ComboBox4.Text = "Juli"
        ElseIf bulan = 8 Then
            ComboBox4.Text = "Agustus"
        ElseIf bulan = 9 Then
            ComboBox4.Text = "September"
        ElseIf bulan = 10 Then
            ComboBox4.Text = "Oktober"
        ElseIf bulan = 11 Then
            ComboBox4.Text = "November"
        ElseIf bulan = 12 Then
            ComboBox4.Text = "Desember"
        End If
        ComboBox5.Text = "20000"
        TextBox8.Text = "311710"
    End Sub
    Sub tampilpengeluaranlain()
        koneksidata.Close()
        Dim ds As DataSet = New DataSet
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        Dim perintah As String = "Select * from pengeluaranlain"
        Dim search As New OleDbDataAdapter(perintah, koneksidata)
        search.Fill(ds, "pengeluaranlain")
        DataGridView5.DataSource = ds.Tables("pengeluaranlain")
        koneksidata.Close()
    End Sub
    Private Sub TabPage4_Enter(sender As Object, e As EventArgs) Handles TabPage4.Enter
        ComboBox9.Items.Clear()
        ComboBox10.Items.Clear()
        bacadataanggota()
        Call tampilpemasukanlain()
        pemasukanlain = True
        hutang = False
        pengeluaranlain = False
        kasbulanan = False
        Button5.Enabled = True
    End Sub
    Sub numotomatispengeluaranlain()
        koneksidata.Close()
        Dim ds As New DataTable
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        Dim perintah As String = "select id_pengeluaranlain from pengeluaranlain order by id_pengeluaranlain desc"
        Dim search As New OleDbDataAdapter(perintah, koneksidata)
        search.Fill(ds)
        If ds.Rows.Count > 0 Then

            TextBox1.Text = ds.Rows(0).Item(0) + 1
        Else

            TextBox1.Text = "4000"
        End If
        koneksidata.Close()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If ComboBox12.Text = "" Then
            MsgBox("ID tidak ada", MsgBoxStyle.Critical, "pesan")

        Else
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "Select * from hutang where id_anggota ='" & ComboBox12.Text & "'"
            perintahdata.Connection = koneksidata
            bacadata = perintahdata.ExecuteReader

            If bacadata.Read = True Then
                pndh2 = bacadata.Item(1).ToString
                pndh3 = bacadata.Item(2).ToString
                pndh4 = bacadata.Item(3).ToString
                ComboBox7.Text = pndh2
                DateTimePicker3.Text = pndh3
                TextBox6.Text = pndh4
                ComboBox12.Enabled = False
                ComboBox7.Enabled = False
                DateTimePicker3.Enabled = False
            Else
                MsgBox("Sodara dengan ID " + ComboBox12.Text + " bernama " + ComboBox7.Text + " tidak memiliki hutang", MsgBoxStyle.Information, "info")
                bacadata.Close()

            End If
            koneksidata.Close()
        End If
        CheckBox1.Enabled = True
    End Sub
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        'If TextBox6.Text.Trim() <> "" Then
        'TextBox6.Text = CDec(TextBox6.Text).ToString("N0")
        'TextBox6.SelectionStart = TextBox6.TextLength
        '  End If
        Label25.Text = pndh4
        Label25.Visible = True

    End Sub
    Sub tampilpemasukanlain()
        koneksidata.Close()
        Dim ds As DataSet = New DataSet
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        Dim perintah As String = "Select * from pemasukanlain"
        Dim search As New OleDbDataAdapter(perintah, koneksidata)
        search.Fill(ds, "pemasukanlain")
        DataGridView3.DataSource = ds.Tables("pemasukanlain")
        koneksidata.Close()
    End Sub
    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter

        Call tampilbulanan()

    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If TextBox6.Text = "" Then
            CheckBox1.Checked = False
        Else

            If CheckBox1.Checked = True Then
                nominal = TextBox6.Text * -1
                TextBox6.Text = nominal
                TextBox6.Enabled = False
            Else
                TextBox6.Text = Label25.Text
                TextBox6.Enabled = True
            End If
        End If
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

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        numotomatispengeluaranlain()
    End Sub
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        'If TextBox7.Text.Trim() <> "" Then
        'TextBox7.Text = CDec(TextBox7.Text).ToString("N0")
        'TextBox7.SelectionStart = TextBox7.TextLength
        ' End If
    End Sub
    Private Sub PrintDataAnggotaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintDataAnggotaToolStripMenuItem.Click
        Form_laporan_anggota.Show()
        Me.Hide()
    End Sub
    Sub cancelhutang()
        TextBox6.Text = ""
        ComboBox7.Text = ""

        ComboBox12.Text = ""

        Button5.Enabled = True

        ComboBox12.Enabled = True
        ComboBox7.Enabled = True
        DateTimePicker3.Enabled = True
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If hutang = True Then
            cancelhutang()
            Label25.Visible = False
            TextBox6.Enabled = True
            CheckBox1.Checked = False
        ElseIf pengeluaranlain = True
            batalpengeluaranlain()

        ElseIf pemasukanlain = True
            batalpemasukanlain()

            Button3.Enabled = True
        ElseIf kasbulanan = True
            batalbulanan()
            ComboBox5.Text = ""
        End If
    End Sub
    Sub hapuspemasuklain()
        Try
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "delete from pemasukanlain where id_pemasukanlain='" & TextBox2.Text & "'"
            perintahdata.Connection = koneksidata
            Dim pertanyaan As Integer = MsgBox("Apakah anda ingin menghapus data ini? ", MsgBoxStyle.YesNo, "Pesan")
            If pertanyaan = DialogResult.Yes Then
                perintahdata.ExecuteNonQuery()
                MsgBox("Berhasil Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
                Call tampilpemasukanlain()

            Else
                MsgBox("Gagal Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
            End If
        Catch ex As Exception
            MsgBox("Gagal menghapus data", MsgBoxStyle.Critical, "Kesalahan")
        End Try
        TextBox3.Enabled = True
        batalpemasukanlain()
    End Sub
    Sub batalpemasukanlain()
        TextBox2.Text = ""
        TextBox2.Enabled = True
        TextBox4.Text = ""
        TextBox4.Enabled = True
        DateTimePicker2.Enabled = True
        TextBox3.Text = ""
        TextBox3.Enabled = True
        ComboBox9.Text = ""
        ComboBox10.Text = ""

    End Sub
    Private Sub PrintLaporanPengeluaranToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintLaporanPengeluaranToolStripMenuItem.Click
        Form_laporan_pengeluaran.Show()
        Me.Hide()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        koneksidata.Close()
        Dim ds As New DataTable
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        Dim perintah As String = "select id_pemasukanlain from pemasukanlain order by id_pemasukanlain desc"
        Dim search As New OleDbDataAdapter(perintah, koneksidata)
        search.Fill(ds)
        If ds.Rows.Count > 0 Then

            TextBox2.Text = ds.Rows(0).Item(0) + 1
        Else

            TextBox2.Text = "1"
        End If
        koneksidata.Close()
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        'If TextBox3.Text.Trim() <> "" Then
        ' TextBox3.Text = CDec(TextBox3.Text).ToString("N0")
        ' TextBox3.SelectionStart = TextBox3.TextLength
        '   End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        cmb1 = False
        cmb2 = False
        cmb3 = False
        cmb4 = False
        cmb5 = True
        cmb6 = False
        cariidanggota()
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        cmb1 = False
        cmb2 = False
        cmb3 = False
        cmb4 = False
        cmb5 = False
        cmb6 = True
        cariidanggota()
    End Sub
    Private Sub TextBox3_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        '  If TextBox3.Text.Trim() <> "" Then
        ' TextBox3.Text = CDec(TextBox3.Text).ToString("N0")
        ' TextBox3.SelectionStart = TextBox3.TextLength
        ' End If
    End Sub
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        koneksidata.Close()
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        perintahdata.CommandText = "Select * from anggota where nama='" & ComboBox7.Text & "'"
        perintahdata.Connection = koneksidata
        bacadata = perintahdata.ExecuteReader

        If bacadata.Read = True Then
            a = bacadata.Item(0).ToString
            b = bacadata.Item(1).ToString
            ComboBox12.Text = a
        Else
            MsgBox("Gagal memuat data", MsgBoxStyle.Information, "info")
            bacadata.Close()

        End If
        koneksidata.Close()
    End Sub
    Private Sub ComboBox9_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        cmb1 = False
        cmb2 = False
        cmb3 = True
        cmb4 = False
        cmb5 = False
        cmb6 = False
        cariidanggota()
    End Sub
    Private Sub ComboBox10_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        cmb1 = False
        cmb2 = False
        cmb3 = False
        cmb4 = True
        cmb5 = False
        cmb6 = False
        cariidanggota()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.CheckState = CheckState.Checked Then
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
        ElseIf CheckBox2.CheckState = CheckState.Unchecked Then
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
        End If
    End Sub

    Private Sub TabPage7_Click(sender As Object, e As EventArgs) Handles TabPage7.Click

    End Sub

    Private Sub AdminToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AdminToolStripMenuItem.Click
        admin.Show()
        Me.Hide()
    End Sub

    Private Sub LogOutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogOutToolStripMenuItem.Click
       end
    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        koneksidata.Close()
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        perintahdata.CommandText = "Select * from anggota where id='" & ComboBox12.Text & "'"
        perintahdata.Connection = koneksidata
        bacadata = perintahdata.ExecuteReader

        If bacadata.Read = True Then
            a = bacadata.Item(0).ToString
            b = bacadata.Item(1).ToString
            ComboBox7.Text = b
        Else

            bacadata.Close()

        End If
        koneksidata.Close()
    End Sub



    Private Sub DataGridView5_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellDoubleClick
        Dim index As Integer
        index = e.RowIndex
        Dim pilihrow As DataGridViewRow
        pilihrow = DataGridView5.Rows(index)

        TextBox1.Text = pilihrow.Cells(0).Value.ToString
        ComboBox6.Text = pilihrow.Cells(1).Value.ToString
        ComboBox8.Text = pilihrow.Cells(2).Value.ToString
        TextBox5.Text = pilihrow.Cells(3).Value.ToString
        DateTimePicker4.Text = pilihrow.Cells(4).Value.ToString
        TextBox7.Text = pilihrow.Cells(5).Value.ToString

        Button4.Enabled = False
        TextBox1.Enabled = False
        TextBox5.Enabled = True
        DateTimePicker4.Enabled = True
    End Sub

    Private Sub TransaksiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransaksiToolStripMenuItem.Click
        TabControl1.Visible = True
    End Sub

    Private Sub KasWajibBulananToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KasWajibBulananToolStripMenuItem.Click
        form_laporan_bulanan.Show()
        Me.Hide()
    End Sub

    Private Sub PemasukanLainToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PemasukanLainToolStripMenuItem.Click
        Form_laporan_pemasukan.Show()
        Me.Hide()
    End Sub
    Private Sub PrintLaporanSaldoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        formlaporansaldo.Show()
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub TabPage5_Click(sender As Object, e As EventArgs) Handles TabPage5.Click

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        cmb1 = True
        cmb2 = False
        cmb3 = False
        cmb4 = False
        cmb5 = False
        cmb6 = False
        cariidanggota()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        cariidtok()
    End Sub


    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        cmb1 = False
        cmb2 = False
        cmb3 = True
        cmb4 = False
        cmb5 = False
        cmb6 = False
        cariidanggota()
    End Sub

    Private Sub PrintLaporanSaldoToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles PrintLaporanSaldoToolStripMenuItem.Click
        Me.Hide()
        form_laporansaldonya.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim controltime As DateTime = Today
        Dim tahun As Integer
        tahun = controltime.Year
        If ComboBox1.Text = tahun Then
        Else
            Dim pertanyaan As Integer = MsgBox("Sekarang Sudah bukan tahun " + ComboBox1.Text + " Data Akan Dikosongkan !, Saya Sarankan Print Data Terlebih Dahulu. ", MsgBoxStyle.YesNo, " Pesan")
            If pertanyaan = DialogResult.Yes Then
                form_laporansaldonya.Show()
                Form_laporan_anggota.Show()
                Form_laporan_pengeluaran.Show()
                form_laporan_bulanan.Show()
                Form_laporan_hutang.Show()
                hapusresetdt()
            End If
        End If
    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        cmb1 = False
        cmb2 = False
        cmb3 = False
        cmb4 = True
        cmb5 = False
        cmb6 = False
        cariidanggota()
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        hapusresetdt()
        tampilbulanan()
        tampilhutang()
        tampilpemasukanlain()
        tampilpengeluaranlain()
        saldo()
    End Sub

    Private Sub PrintLaporanHutangToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintLaporanHutangToolStripMenuItem.Click
        Form_laporan_hutang.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)
        cariidtok()
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        cmb1 = False
        cmb2 = True
        cmb3 = False
        cmb4 = False
        cmb5 = False
        cmb6 = False
        cariidanggota()
    End Sub
    Sub cariidtok()
        koneksidata.Close()
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        If cmb1 = True Then
            perintahdata.CommandText = "Select * from anggota where id='" & ComboBox6.Text & "'"
        ElseIf cmb3 = True
            perintahdata.CommandText = "Select * from anggota where id='" & ComboBox9.Text & "'"
        ElseIf cmb5 = True
            perintahdata.CommandText = "Select * from anggota where id='" & ComboBox2.Text & "'"
        End If
        perintahdata.Connection = koneksidata
        bacadata = perintahdata.ExecuteReader

        If bacadata.Read = True Then
            a = bacadata.Item(0).ToString
            b = bacadata.Item(1).ToString
            ComboBox6.Text = a
            ComboBox9.Text = a

            If pengeluaranlain = True Then
                If ComboBox6.Text = a Then
                    ComboBox8.Text = b
                End If
            ElseIf pemasukanlain = True
                If ComboBox9.Text = a Then
                    ComboBox10.Text = b
                End If
            ElseIf kasbulanan = True
                If ComboBox2.Text = a Then
                    ComboBox3.Text = b
                End If
            End If
        Else

            bacadata.Close()

        End If
        bacadata.Close()

        koneksidata.Close()
    End Sub
    Sub carinim()
        koneksidata.Close()
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        perintahdata.CommandText = "Select * from bulanan where id_anggota='" & TextBox8.Text & "'"

    End Sub

    Sub cariidanggota()
        koneksidata.Close()
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        If cmb1 = True Then
            perintahdata.CommandText = "Select * from anggota where id='" & ComboBox6.Text & "'"
        ElseIf cmb2 = True
            perintahdata.CommandText = "Select * from anggota where nama='" & ComboBox8.Text & "'"
        ElseIf cmb3 = True
            perintahdata.CommandText = "Select * from anggota where id='" & ComboBox9.Text & "'"
        ElseIf cmb4 = True
            perintahdata.CommandText = "Select * from anggota where nama='" & ComboBox10.Text & "'"
        ElseIf cmb5 = True
            perintahdata.CommandText = "Select * from anggota where id='" & ComboBox2.Text & "'"
        ElseIf cmb6 = True
            perintahdata.CommandText = "Select * from anggota where nama='" & ComboBox3.Text & "'"
        End If
        perintahdata.Connection = koneksidata
        bacadata = perintahdata.ExecuteReader

        If bacadata.Read = True Then
            a = bacadata.Item(0).ToString
            b = bacadata.Item(1).ToString
            If pengeluaranlain = True Then
                ComboBox8.Text = b
                ComboBox6.Text = a
            ElseIf pemasukanlain = True
                ComboBox9.Text = a
                ComboBox10.Text = b
            ElseIf kasbulanan = True
                ComboBox2.Text = a
                ComboBox3.Text = b
            End If
        Else

            bacadata.Close()

        End If
        bacadata.Close()

        koneksidata.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        cariidtok()
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress

        If (e.KeyChar < "0" OrElse e.KeyChar > "9") _
           AndAlso e.KeyChar <> ControlChars.Back Then 'AndAlso e.KeyChar <> "-"
            e.Handled = True
        End If
    End Sub
    Private Sub DataGridView3_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellDoubleClick
        Dim index As Integer
        index = e.RowIndex
        Dim pilihrow As DataGridViewRow
        pilihrow = DataGridView3.Rows(index)

        TextBox2.Text = pilihrow.Cells(0).Value.ToString
        ComboBox9.Text = pilihrow.Cells(1).Value.ToString
        ComboBox10.Text = pilihrow.Cells(2).Value.ToString
        TextBox4.Text = pilihrow.Cells(3).Value.ToString
        DateTimePicker2.Text = pilihrow.Cells(4).Value.ToString
        TextBox3.Text = pilihrow.Cells(5).Value.ToString

        Button3.Enabled = False

        TextBox2.Enabled = False
        TextBox4.Enabled = True
        DateTimePicker2.Enabled = True


    End Sub
    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        ComboBox12.Items.Clear()
        ComboBox7.Items.Clear()
        bacadataanggota()
        Call tampilhutang()
        hutang = True
        pengeluaranlain = False
        pemasukanlain = False
    End Sub
    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If (e.KeyChar < "0" OrElse e.KeyChar > "9") _
       AndAlso e.KeyChar <> ControlChars.Back Then 'AndAlso e.KeyChar <> "-"
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If (e.KeyChar < "0" OrElse e.KeyChar > "9") _
       AndAlso e.KeyChar <> ControlChars.Back Then 'AndAlso e.KeyChar <> "-"
            e.Handled = True
        End If
    End Sub
    Private Sub ComboBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox5.KeyPress
        If (e.KeyChar < "0" OrElse e.KeyChar > "9") _
     AndAlso e.KeyChar <> ControlChars.Back Then 'AndAlso e.KeyChar <> "-"
            e.Handled = True
        End If
    End Sub
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim index As Integer
        index = e.RowIndex
        Dim pilihrow As DataGridViewRow
        pilihrow = DataGridView1.Rows(index)

        ComboBox2.Text = pilihrow.Cells(0).Value.ToString
        ComboBox3.Text = pilihrow.Cells(1).Value.ToString
        ComboBox1.Text = pilihrow.Cells(2).Value.ToString
        If ComboBox4.Text = "Januari" Then
            ComboBox5.Text = pilihrow.Cells(3).Value.ToString
        ElseIf ComboBox4.Text = "Februari" Then
            ComboBox5.Text = pilihrow.Cells(4).Value.ToString
        ElseIf ComboBox4.Text = "Maret" Then
            ComboBox5.Text = pilihrow.Cells(5).Value.ToString
        ElseIf ComboBox4.Text = "April" Then
            ComboBox5.Text = pilihrow.Cells(6).Value.ToString
        ElseIf ComboBox4.Text = "Mei" Then
            ComboBox5.Text = pilihrow.Cells(7).Value.ToString
        ElseIf ComboBox4.Text = "Juni" Then
            ComboBox5.Text = pilihrow.Cells(8).Value.ToString
        ElseIf ComboBox4.Text = "Juli" Then
            ComboBox5.Text = pilihrow.Cells(9).Value.ToString
        ElseIf ComboBox4.Text = "Agustus" Then
            ComboBox5.Text = pilihrow.Cells(10).Value.ToString
        ElseIf ComboBox4.Text = "September" Then
            ComboBox5.Text = pilihrow.Cells(11).Value.ToString
        ElseIf ComboBox4.Text = "Oktober" Then
            ComboBox5.Text = pilihrow.Cells(12).Value.ToString
        ElseIf ComboBox4.Text = "November" Then
            ComboBox5.Text = pilihrow.Cells(13).Value.ToString
        ElseIf ComboBox4.Text = "Desember" Then
            ComboBox5.Text = pilihrow.Cells(14).Value.ToString
        End If
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        End
    End Sub

    Private Sub ComboBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox2.KeyPress

    End Sub

    Private Sub TabPage5_Enter(sender As Object, e As EventArgs) Handles TabPage5.Enter
        saldong()
        saldo()
    End Sub
    Sub saldo()
        koneksidata.Close()
        Dim ds As DataSet = New DataSet
        koneksidata.ConnectionString = koneksistring
        koneksidata.Open()
        Dim perintah As String = "Select * from saldo"
        Dim search As New OleDbDataAdapter(perintah, koneksidata)
        search.Fill(ds, "saldo")
        DataGridView2.DataSource = ds.Tables("saldo")
        koneksidata.Close()
    End Sub
    Sub saldong()

        Dim januari As Integer
        januari = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            januari = januari + Val(DataGridView1.Rows(t).Cells(3).Value)
        Next
        Dim feb As Integer
        feb = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            feb = feb + Val(DataGridView1.Rows(t).Cells(4).Value)
        Next
        Dim mar As Integer
        mar = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            mar = mar + Val(DataGridView1.Rows(t).Cells(5).Value)
        Next
        Dim apr As Integer
        apr = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            apr = apr + Val(DataGridView1.Rows(t).Cells(6).Value)
        Next
        Dim mei As Integer
        mei = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            mei = mei + Val(DataGridView1.Rows(t).Cells(7).Value)
        Next
        Dim jun As Integer
        jun = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            jun = jun + Val(DataGridView1.Rows(t).Cells(8).Value)
        Next
        Dim jul As Integer
        jul = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            jul = jul + Val(DataGridView1.Rows(t).Cells(9).Value)
        Next
        Dim agu As Integer
        agu = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            agu = agu + Val(DataGridView1.Rows(t).Cells(10).Value)
        Next
        Dim sep As Integer
        sep = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            sep = sep + Val(DataGridView1.Rows(t).Cells(11).Value)
        Next
        Dim okt As Integer
        okt = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            okt = okt + Val(DataGridView1.Rows(t).Cells(12).Value)
        Next
        Dim nov As Integer
        nov = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            nov = nov + Val(DataGridView1.Rows(t).Cells(13).Value)
        Next
        Dim des As Integer
        des = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            des = des + Val(DataGridView1.Rows(t).Cells(14).Value)
        Next
        Dim pem As Integer
        pem = 0
        For t As Integer = 0 To DataGridView3.Rows.Count - 1
            pem = pem + Val(DataGridView3.Rows(t).Cells(5).Value)
        Next
        Dim peng As Integer
        peng = 0
        For t As Integer = 0 To DataGridView5.Rows.Count - 1
            peng = peng + Val(DataGridView5.Rows(t).Cells(5).Value)
        Next
        Dim hut As Integer
        hut = 0
        For t As Integer = 0 To DataGridView4.Rows.Count - 1
            hut = hut + Val(DataGridView4.Rows(t).Cells(3).Value)
        Next
        Dim satu As Integer = 1
        Try
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "insert into saldo (kode,jan,feb,mar,apr,mei,jun,jul,ags,sep,okt,nov,des,pem,peng,hut) values ('" & satu & "','" & januari & "','" & feb & "','" & mar & "','" & apr & "','" & mei & "','" & jun & "','" & jul & "','" & agu & "','" & sep & "','" & okt & "','" & nov & "','" & des & "','" & pem & "','" & peng & "','" & hut & "')"
            perintahdata.Connection = koneksidata
            perintahdata.ExecuteNonQuery()
            koneksidata.Close()
            perintahdata.Dispose()
            koneksidata.Close()
            Call tampilpengeluaranlain()
        Catch ex As Exception
            Dim eh As Integer = 1
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "update saldo set jan='" & januari & "',feb='" & feb & "', mar='" & mar & "',apr='" & apr & "',mei='" & mei & "',jun='" & jun & "',jul='" & jul & "',ags='" & agu & "',sep ='" & sep & "',okt='" & okt & "',nov='" & nov & "',des='" & des & "',pem='" & pem & "',peng='" & peng & "',hut='" & hut & "' where kode='" & 1 & "'"
            perintahdata.Connection = koneksidata
            perintahdata.ExecuteNonQuery()
            koneksidata.Close()
            perintahdata.Dispose()
        End Try
        saldo()
    End Sub
    Sub hapusresetdt()
        Try
            Dim pertanyaan As Integer = MsgBox("Data Transaksi Akan Hilang Semua, Saya Saran kan Print Laporan Terlebih Dahulu! Hapus Data? ", MsgBoxStyle.YesNo, "Pesan")
            If pertanyaan = DialogResult.Yes Then
                If pertanyaan = DialogResult.Yes Then
                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "delete from bulanan where id_anggota"
                    perintahdata.Connection = koneksidata
                    perintahdata.ExecuteNonQuery()
                    koneksidata.Close()
                    perintahdata.Dispose()
                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "delete from saldo where kode"
                    perintahdata.Connection = koneksidata
                    perintahdata.ExecuteNonQuery()
                    koneksidata.Close()
                    perintahdata.Dispose()
                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "delete from pemasukanlain where id_pemasukanlain"
                    perintahdata.Connection = koneksidata
                    perintahdata.ExecuteNonQuery()
                    koneksidata.Close()
                    perintahdata.Dispose()
                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "delete from hutang where id_anggota"
                    perintahdata.Connection = koneksidata
                    perintahdata.ExecuteNonQuery()
                    koneksidata.Close()
                    perintahdata.Dispose()
                    koneksidata.Close()
                    koneksidata.ConnectionString = koneksistring
                    koneksidata.Open()
                    perintahdata.CommandText = "delete from pengeluaranlain where id_pengeluaranlain"
                    perintahdata.Connection = koneksidata
                    perintahdata.ExecuteNonQuery()
                    MsgBox("Berhasil Menghapus", MsgBoxStyle.Information, "Info")
                    koneksidata.Close()
                    perintahdata.Dispose()
                End If
            Else
                MsgBox("Gagal Menghapus", MsgBoxStyle.Information, "Info")
                koneksidata.Close()
                perintahdata.Dispose()
            End If
        Catch ex As Exception
        End Try
    End Sub

End Class
