Public Class Form_laporan_pemasukan
    Private Sub Form_laporan_pemasukan_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Form1.Show()
    End Sub

    Private Sub Form_laporan_pemasukan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CrystalReportViewer1.RefreshReport()
    End Sub


End Class