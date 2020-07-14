Public Class form_laporan_seluruh_pemasukan
    Private Sub form_laporan_seluruh_pemasukan_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Form1.Show()
    End Sub

    Private Sub form_laporan_seluruh_pemasukan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CrystalReportViewer1.RefreshReport()
    End Sub

    Private Sub seluruh_pemasukan1_InitReport(sender As Object, e As EventArgs) 

    End Sub
End Class