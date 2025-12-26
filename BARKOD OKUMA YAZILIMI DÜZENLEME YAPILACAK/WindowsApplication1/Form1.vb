Imports System.ComponentModel
Imports System.IO
Imports Microsoft.Office.Interop

'--------------------------------------------------------------
' Samet UĞRAÇ 21.05.2019 Tarihinde UYGUR TEKER için yapmıştır.|
'--------------------------------------------------------------


'PROGRAMI OLDUKÇA YALIN BİR BİÇİMDE ANLATTIM

Public Class Form1
    '--------------------------------------------------------------------------------------------------------------------------------------------------
    'GLOBAL DEĞİŞKENLERİMİZ HER YERDEN ERİŞEBİLELİM DİYE
    Dim x
    Dim a
    '--------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.RowsDefaultCellStyle.BackColor = Color.Yellow

        TextBox2.TextAlign = HorizontalAlignment.Center
        TextBox1.Select()
        TextBox2.Visible = True
        Timer2.Enabled = True
        TextBox2.Text = 0
        'BAZI YERLERDE ---- TEXTBOX1.SELECT ---- GÖREBİLİRSİNİZ BU O KOD BLOĞUNA BAĞLI NESNEYE TIKLANINCA TEXTBOX1.FOCUS'LAN ANLANMINDADIR.


    End Sub


    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        '--------------------------------------------------------------------------------------------------------------------------------------------------
        'BURASIDA TEXT1 OTOMATİK CLİCK YAPAN BUTONUMUZ 1 SANİYEDE BİR TIKLIYOR BUTONA 
        Button1.PerformClick()
        '--------------------------------------------------------------------------------------------------------------------------------------------------

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '--------------------------------------------------------------------------------------------------------------------------------------------------


        'FORM İÇİNDE GİZLİ BUTTON VAR BU BUTON ÜZERİNDEN KULLANICI YORULMASIN DİYE  OTOMATİK TIKLAMA İÇİN OLUŞTURULDU
        If TextBox1.Text = "" Then
            Timer1.Enabled = False
            Timer2.Enabled = True

        Else
            '--------------------------------------------------------------------------------------------------------------------------------------------------

            'TEXTBOX2 SAĞ TARAFTA GÖRÜNEN  KÜÇÜK SARI KUTU BU TEXTBOX SAYAC GÖREVİ GÖRÜYOR HER NE KADAR DOĞRU ÇALIŞMASA'DA

            Timer1.Enabled = True
            'İŞTE ZURNANIN ZIRT DEDİĞİ YER 2 GÜN BOYUNCA 2 SATIR İÇİN UĞRAŞTIM
            'BURADA TEXT1 GİRİLEN DEĞER SAYISINI  UZUNLIĞUNU A DEĞİŞKENİNE AKTARIYORUZ
            'SONRA DATAGRİD EKRANA SATIR EKLERKEN DİYORU Kİ TEXT2(SIRA NO)  TEXT1(BARKOD NUMARASI)  A(SAYAC) SATIRLARINI DATAGRİD EKLİYOR
            Dim a = TextBox1.Text.Length
            DataGridView1.Rows.Add(TextBox2.Text, TextBox1.Text, a)

            '--------------------------------------------------------------------------------------------------------------------------------------------------


            'GİRİLEN SON DEĞERİ LABEL GÖSTERMEK İÇİN TEXT1 İLE EŞİTLEDİM
            Label1.Text = TextBox1.Text


            '---------------------------------------------------------------------------------------------------------------------------------
            For x = 0 To DataGridView1.Rows.Count - 2 Step 1  'for döngüsü devreye giriyor
                DataGridView1.Rows(x).DefaultCellStyle.BackColor = Color.Pink
                TextBox2.Clear()
                TextBox1.Clear()
                TextBox1.Select()

                'DATAGRİD ÜZERİNDE BULUNAN VERİLER DEVAMLI SCOLLBAR YARDIMI İLE AŞAĞI ÇEKİLİYOR GİRİLEN SON VERİ GÖZÜKÜYOR
                DataGridView1.FirstDisplayedScrollingRowIndex = x
            Next
            TextBox2.Text = x
        End If
        '---------------------------------------------------------------------------------------------------------------------------------
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        '---------------------------------------------------------------------------------------------------------------------------------
        'NE İÇİN YAPTIĞIMI HATIRLAMIYORUM
        TextBox1.Select()
        Dim satir As Integer
        satir = DataGridView1.CurrentRow.Index
        '---------------------------------------------------------------------------------------------------------------------------------

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox1.Select()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        '-----------------------------------//////BURASI ÇOK ÖNEMLİ//////-----------------------------------------------

        'excel referans ekleme için
        'solution exploer üzerinde bulunan Rferances gelip sağ tıklıyoruz
        'Add referances diyoruz
        'COM bulunan referandan microsoft interop.excel arıyoruz  
        'ctrl+f yardımı ile excel yazarsak hemen bulunur

        '---------------------------------------------------------------------------------------------------------------------------------


        'C:\\ klasörünün içinde Bilgi_Sayar isimli bir klasör var ise Boolean türündeki değişkenin içine True yazdırılır. Yoksa False yazdırılır.
        Dim Belgelerim As String
        Belgelerim = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        Dim klasor_varmi As Boolean = My.Computer.FileSystem.DirectoryExists(Belgelerim + ".\Çıktılar")

        '---------------------------------------------------------------------------------------------------------------------------------
        If klasor_varmi = False Then
            Directory.CreateDirectory(Belgelerim + ".\Çıktılar")

        End If

        '---------------------------------------------------------------------------------------------------------------------------------

        'DATAGRİD ÜZERİNE OLUŞTURDUĞUMUZ VERİLERİ EXCELE AKTARMASI İÇİN KULLANDIĞIMIZ DEĞİŞKENLER
        Dim excel As Excel.Application
            Dim kitap As Excel.Workbook
            Dim sayfa As Excel.Worksheet
            Dim datagrid As DataGridView
            datagrid = Me.DataGridView1
            excel = CreateObject("Excel.Application")
            kitap = excel.Workbooks.Add
            sayfa = excel.Worksheets(1)

            '---------------------------------------------------------------------------------------------------------------------------------


            '-------------------EXCELDEKİ  HEADER KISIMLARI İÇİN AYRILAN BÖLÜM BURADA YAZI TİPİ VE SUTUN GENİŞLİĞİNİ AYARLIYORUZ---------------------------


            'exceldeki header bölümü yazı tipleri
            excel.Application.Cells(1, 1).Font.Size = 15
            excel.Application.Cells(1, 2).Font.Size = 15
            excel.Application.Cells(1, 3).Font.Size = 15

            '---------------------------------------------------------------------------------------------------------------------------------

            'Yazı kalınlıkları
            excel.Application.Cells(1, 1).Font.Bold = True
            excel.Application.Cells(1, 2).Font.Bold = True
            excel.Application.Cells(1, 3).Font.Bold = True

            '---------------------------------------------------------------------------------------------------------------------------------

            'Excel Kolon Genişliği
            excel.Application.Cells(1, 1).ColumnWidth = 10
            excel.Application.Cells(1, 2).ColumnWidth = 34
            excel.Application.Cells(1, 3).ColumnWidth = 26

            '---------------------------------------------------------------------------------------------------------------------------------


            If klasor_varmi = True Then


                '---------------------------------------------------------------------------------------------------------------------------------
                'FOR DÖNGÜSÜ BURADA DATAGRİD ÜZERİNDEN EXCELE BİLGİ AKTARIYORUZ
                For yaz = 1 To datagrid.RowCount
                    sayfa.Cells(yaz, 1).value = datagrid.Rows(yaz - 1).Cells(0).Value
                    sayfa.Cells(yaz, 2).value = datagrid.Rows(yaz - 1).Cells(1).Value
                    sayfa.Cells(yaz, 3).value = datagrid.Rows(yaz - 1).Cells(2).Value


                    ' x(açıklaması)= Excel üzerinde hangi columns yazıyım diyor
                    '------------------------------------------------
                    'sayfa.Cells(yaz,X).value =
                    '------------------------------------------------


                    'x(açıklası)=datagriddeki bilgileri HANGİ HÜCRE BİLGİSİNE göre çekiyim diye soruyor saymaya 0'dan başlıyor dikkat et
                    '------------------------------------------------
                    ' datagrid.Rows(yaz - 1).Cells(X).Value
                    '------------------------------------------------
                Next



                'EXCEL HEADER İSİMLERİNİ BURADA GÖSTERTİYORUZ
                sayfa.Cells(1, 1).value = "Satır No"
                sayfa.Cells(1, 2).value = "Barkod Numarası"
                sayfa.Cells(1, 3).value = "Okunan Değer Sayısı"


                '---------------------------------------------------------------------------------------------------------------------------------
                'Excel dosyalarını Gün gün kayıt etmek için burda tarih tanımladık
                Dim tarih2 As Date
                tarih2 = DateTime.Now.ToShortDateString()

                '--------------------------------------------------------------------------------------------------------------------------------------------------

                'burada  belgelerim klasörü içinde  çıktılar klasörü kurulacak onun için sev edecek program
                kitap.SaveAs(".\Çıktılar\" + tarih2 + ".xlsx")

                excel.Quit()
                MsgBox("Belgerim\Çıktılar\ Klasörüne Excel Formatında Aktarıldı")
            End If


        TextBox1.Select()

    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        TextBox1.Select()

    End Sub

    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click
        TextBox1.Select()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox1.Select()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Process.Start("C:\Users\SAMET\Documents\Çıktılar\")



    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub
End Class

'----------------------------------------------------------------------------------------...more