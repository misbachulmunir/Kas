Public Class Form_laporan_anggota
    Private Sub Form_laporan_anggota_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Form1.Show()
    End Sub

    Private Sub Form_laporan_anggota_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CrystalReportViewer1.RefreshReport()
    End Sub
End Class