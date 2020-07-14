Imports System.Data.OleDb

Public Class login
    Dim koneksistring As String
    Dim koneksidata As New OleDbConnection
    Dim perintahdata As New OleDbCommand
    Dim adapter As New OleDbDataAdapter
    Dim table As New DataTable
    Dim bacadata As OleDbDataReader
    Dim path As String = My.Application.Info.DirectoryPath + "\"
    Dim ac, bc As String
    Private Sub login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = 0
        Dim b As New Drawing2D.GraphicsPath
        b.AddEllipse(0, 0, Me.Width, Me.Height)
        Me.Region = New Region(b)
        koneksistring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & "kas.mdb"
        Try
            koneksidata.Close()
            koneksidata.ConnectionString = koneksistring
            koneksidata.Open()
            perintahdata.CommandText = "Select * from admin"
            perintahdata.Connection = koneksidata
            bacadata = perintahdata.ExecuteReader
            If bacadata.Read = True Then

                Do
                    ac = bacadata.Item(1).ToString
                    bc = bacadata.Item(2).ToString


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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = ac And TextBox2.Text = bc Then
            Form1.Show()
            Me.Hide()
        Else
            MsgBox("Username atau Password anda salah !", MsgBoxStyle.Information, "Info")
        End If
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub


    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub
End Class