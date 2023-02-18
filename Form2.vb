Imports Word = Microsoft.Office.Interop.Word

Imports System.IO
Imports System.Threading

Public Class Form2
    Dim App As New Word.Application
    Dim Document As Word.Document
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For a = 1 To 31
            cmbHari.Items.Add(a)
        Next
        For c = 2023 To 2015 Step -1
            cmbTahun.Items.Add(c)
        Next
    End Sub

    Private Sub cmbMenuMakanan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMenuMakanan.SelectedIndexChanged
        Select Case cmbMenuMakanan.Text
            Case "Sandwich"
                lblhargamkn.Text = "19000"
                PictureBox2.Image = System.Drawing.Image.FromFile("C:\Users\PERSONAL\Downloads\sabrina\Pemrograman Visual\sandwich.jpg")
            Case "Spagheti"
                lblhargamkn.Text = "20000"
                PictureBox2.Image = System.Drawing.Image.FromFile("C:\Users\PERSONAL\Downloads\sabrina\Pemrograman Visual\spagheti1.jpg")
            Case "Burger"
                lblhargamkn.Text = "15000"
                PictureBox2.Image = System.Drawing.Image.FromFile("C:\Users\PERSONAL\Downloads\sabrina\Pemrograman Visual\burger1.jpg")
            Case "Chicken Rice"
                lblhargamkn.Text = "23000"
                PictureBox2.Image = System.Drawing.Image.FromFile("C:\Users\PERSONAL\Downloads\sabrina\Pemrograman Visual\rice1.jpg")
            Case "Salad"
                lblhargamkn.Text = "17000"
                PictureBox2.Image = System.Drawing.Image.FromFile("C:\Users\PERSONAL\Downloads\sabrina\Pemrograman Visual\salad1.jpg")
        End Select
    End Sub

    Private Sub cmbMenuMinuman_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMenuMinuman.SelectedIndexChanged
        Select Case cmbMenuMinuman.Text
            Case "Manggo Juice"
                lblhrgminum.Text = "10000"
                PictureBox3.Image = System.Drawing.Image.FromFile("C:\Users\PERSONAL\Downloads\sabrina\Pemrograman Visual\Mango1.jpg")
            Case "Blueberry Juice"
                lblhrgminum.Text = "15000"
                PictureBox3.Image = System.Drawing.Image.FromFile("C:\Users\PERSONAL\Downloads\sabrina\Pemrograman Visual\blueberry1.jpg")
            Case "Choco Milky"
                lblhrgminum.Text = "19000"
                PictureBox3.Image = System.Drawing.Image.FromFile("C:\Users\PERSONAL\Downloads\sabrina\Pemrograman Visual\choco1.jpg")
            Case "Grean Tea"
                lblhrgminum.Text = "10000"
                PictureBox3.Image = System.Drawing.Image.FromFile("C:\Users\PERSONAL\Downloads\sabrina\Pemrograman Visual\grean1.jpg")
            Case "Ice Cream"
                lblhrgminum.Text = "13000"
                PictureBox3.Image = System.Drawing.Image.FromFile("C:\Users\PERSONAL\Downloads\sabrina\Pemrograman Visual\ice1.jpg")
        End Select
    End Sub

    Private Sub DessertPicture1_Click(sender As Object, e As EventArgs) Handles DessertPicture1.Click
        nameDessrt.Text = "Cake Fruit"
        hrgdsrt.Text = "12000"
    End Sub
    Private Sub DessertPicture2_Click(sender As Object, e As EventArgs) Handles DessertPicture2.Click
        hrgdsrt.Text = "13000"
    End Sub
    Private Sub DessertPicture3_Click(sender As Object, e As EventArgs) Handles DessertPicture3.Click
        hrgdsrt.Text = "22000"
    End Sub
    Private Sub DessertPicture4_Click(sender As Object, e As EventArgs) Handles DessertPicture4.Click
        hrgdsrt.Text = "25000"
    End Sub

    Private Sub btnHitung_Click(sender As Object, e As EventArgs) Handles btnHitung.Click
        txtTotal.Text = (lblhargamkn.Text * jmlhmknan.Text) + (lblhrgminum.Text * jmlhminuman.Text) + (hrgdsrt.Text * txtJmlhDessert.Text)
        Label18.Text = Format(Now, "dd MMMM yyyy")
        SabrinaCaffe.Visible = True
        SabrinaCaffe.Items.Add("SABRINA CAFFE")
        SabrinaCaffe.Items.Add("Nama Pembeli : " + txtNama.Text)
        SabrinaCaffe.Items.Add("No Meja : " + txtMeja.Text)
        SabrinaCaffe.Items.Add("Tanggal : " + Label18.Text)
        SabrinaCaffe.Items.Add("Makanan : " + cmbMenuMakanan.SelectedItem)
        SabrinaCaffe.Items.Add("Jumlah Makanan : " + jmlhmknan.Text)
        SabrinaCaffe.Items.Add("Minuman : " + cmbMenuMinuman.SelectedItem)
        SabrinaCaffe.Items.Add("Jumlah Minuman : " + jmlhminuman.Text)
        SabrinaCaffe.Items.Add("Dessert : " + hrgdsrt.Text)
        SabrinaCaffe.Items.Add("Jumlah Dissert : " + txtJmlhDessert.Text)
        SabrinaCaffe.Items.Add("Total Harga : " + txtTotal.Text)
    End Sub

    Private Sub btnPesan_Click(sender As Object, e As EventArgs) Handles btnPesan.Click
        txtKembalian.Text = (txtBayar.Text - txtTotal.Text)
        MessageBox.Show("Terimakasih Pesanan anda segera kami antar", "Sabrina Cafee", MessageBoxButtons.OK, MessageBoxIcon.None)
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim x = MsgBox("Yakin sudah selesai?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Confirmation")
        If x = vbYes Then
            Me.Close()
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Document = App.Documents.Open("C:\Users\PERSONAL\Documents\Custom Office Templates\Sabrina-cafee.dotx")
        Document.Bookmarks("nama").Select()
        App.Selection.TypeText(txtNama.Text)
        Document.Bookmarks("tanggal").Select()
        App.Selection.TypeText(Label18.Text)
        Document.Bookmarks("meja").Select()
        App.Selection.TypeText(txtMeja.Text)
        Document.Bookmarks("makanan").Select()
        App.Selection.TypeText(cmbMenuMakanan.Text)
        Document.Bookmarks("jumlahMakanan").Select()
        App.Selection.TypeText(jmlhmknan.Text)
        Document.Bookmarks("minuman").Select()
        App.Selection.TypeText(cmbMenuMinuman.Text)
        Document.Bookmarks("jumlahMinuman").Select()
        App.Selection.TypeText(jmlhminuman.Text)
        Document.Bookmarks("dessert").Select()
        App.Selection.TypeText(hrgdsrt.Text)
        Document.Bookmarks("jumlahDessert").Select()
        App.Selection.TypeText(txtJmlhDessert.Text)
        Document.Bookmarks("total").Select()
        App.Selection.TypeText(txtTotal.Text)
        Document.Bookmarks("dibayar").Select()
        App.Selection.TypeText(txtBayar.Text)
        Document.Bookmarks("kembalian").Select()
        App.Selection.TypeText(txtKembalian.Text)

        Document.SaveAs2("C:\Users\PERSONAL\Documents\Custom Office Templates\Sabrina-cafee-1.dotx")
        App.Visible = True
    End Sub
End Class