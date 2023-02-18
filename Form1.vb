Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MessageBox.Show("Hello Selamat Datang di Sabrina Cafee", "Sabrina Cafee")
        Form2.Show()
        Me.Hide()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MessageBox.Show("Bismillahirrahmanirrahiim. Sabrina Cafee berdiri sejak 2022 dan selalu ingin memberi hasil yang terbaik bagi pelanggan. Terimakasih ", "Sabrina Cafee")
        Me.Show()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim x = MsgBox("Yakin ingin keluar?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Konfirmasi")
        If x = vbYes Then
            Me.Close()
        End If
    End Sub
End Class
