Public Class Form1

    Private Sub btnMainMenu_Click(sender As Object, e As EventArgs) Handles btnMainMenu.Click
        MessageBox.Show("Hello Selamat Datang di Sabrina Cafee", "Sabrina Cafee")
        Form2.Show()
        Me.Hide()
    End Sub
    Private Sub btnAboutUs_Click(sender As Object, e As EventArgs) Handles btnAboutUs.Click
        MessageBox.Show("Bismillahirrahmanirrahiim. Sabrina Cafee berdiri sejak 2022 dan selalu ingin memberi hasil yang terbaik bagi pelanggan. Terimakasih ", "Sabrina Cafee")
        Me.Show()
    End Sub
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim x = MsgBox("Yakin ingin keluar?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Konfirmasi")
        If x = vbYes Then
            Me.Close()
        End If
    End Sub
End Class
