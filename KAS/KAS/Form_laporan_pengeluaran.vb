Public Class Form_laporan_pengeluaran
    Private Sub Form_laporan_pengeluaran_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Form1.Show()
    End Sub

    Private Sub Form_laporan_pengeluaran_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CrystalReportViewer1.RefreshReport()

    End Sub


End Class