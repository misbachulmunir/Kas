Public Class Form_laporan_hutang
    Private Sub Form_laporan_hutang_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Form1.Show()

    End Sub

    Private Sub Form_laporan_hutang_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.CrystalReportViewer1.RefreshReport()
    End Sub

    Private Sub Report_hutang1_InitReport(sender As Object, e As EventArgs)

    End Sub
End Class