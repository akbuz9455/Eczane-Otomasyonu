using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//sql kullanma kütüphanemizi yükledik
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
namespace EczaneOtomasyonu
{
    public partial class ilacBilgileriForm : Form
    {
        sqlbaglantisi bag = new sqlbaglantisi(); // sql bağlantı classımızdan nesne oluşturduk ve bağlantı için kullanacağız
        DataSet dtst = new DataSet(); //datagridviewlere sql serverdaki tabloları aktarmak için kullanıyoruz.
        SqlCommand kmt = new SqlCommand();  //sql ekleme silme güncelleme listeleme işlemleri için sqlcommand nesnesi oluşturduk
        public ilacBilgileriForm()
        {
            InitializeComponent();
        }


        public void ilacListesiDoldur() //ekleme silme güncelleme işlemlerinden sonra tekrar listelenebilmesi için datagridwiewi doldurma işlemini fonksiyon şeklinde yazdık
        {
            dtst.Clear();

            SqlDataAdapter adtr = new SqlDataAdapter("select * From ilac", bag.baglan());
            adtr.Fill(dtst, "ilac");
            dataGridView1.DataMember = "ilac";
            dataGridView1.DataSource = dtst;
            adtr.Dispose();
            //listenilen ilaçları datagridwiewe aktarır
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Barkod No";
            dataGridView1.Columns[2].HeaderText = "İlacın Adı";
            dataGridView1.Columns[3].HeaderText = "Üretici Firma";
            dataGridView1.Columns[4].HeaderText = "Kutu Sayısı";
            dataGridView1.Columns[5].HeaderText = "Fiyatı";
            dataGridView1.Columns[6].HeaderText = "kullanım Amacı";
            dataGridView1.Columns[7].HeaderText = "Yan Etkileri";
            dataGridView1.Columns[8].HeaderText = "İlacı Teslim Alan Personel";
            //form üzerindeki textbox ve comboboxları temizledik
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox8.Text = "";
            richTextBox1.Text = "";
        }

        private void ilacBilgileriForm_Load(object sender, EventArgs e)
        {//oluşturdumuz fonksiyonları form yüklenirken çağırdık
            ilacListesiDoldur();
        }

        private void button1_Click(object sender, EventArgs e)//ekleme işlemini yapacak butonumuz
        {
           
           kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "INSERT INTO ilac(barkod_no,ilacin_adi,uretici_firma,kutu_sayisi,fiyati,kullanim_amaci,yan_etkileri,ilac_teslim_alan) VALUES ('" + textBox1.Text + "' ,'" + textBox2.Text + "' ,'" + textBox3.Text + "' ,'" + textBox4.Text + "' ,'" + textBox5.Text + "' ,'" + textBox6.Text + "' ,'" + richTextBox1.Text+ "' ,'" + textBox8.Text + "'  )";
            kmt.ExecuteNonQuery();//ekleme sql komutunu çalıştırdık ve başarılı mesaj verip gridwiewi tekrar doldurma fonksiyonunu çağırdık.
            kmt.Dispose();
          
            MessageBox.Show("İlaç kayıt işlemi tamamlandı ! ");
           dtst.Clear();
            ilacListesiDoldur();
        }

        private void button4_Click(object sender, EventArgs e) //bu butonumuzda seçili datagridviewdeki id ye göre silme işlemi yaptık
        {
           
            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "DELETE from ilac WHERE id = '" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'";

           kmt.ExecuteNonQuery(); //oluşturdumuz silme sql komutunu burada çalıştırdık ve aşşağıda başarılı mesajını verip formu tekrar doldurma fonksiyonunu çağırdık
            kmt.Dispose();
            MessageBox.Show("İlaç Silme işlemi tamamlandı ! ");
            dtst.Clear();
         ilacListesiDoldur();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ilacListesiDoldur();//ilaç listesini doldurma fonksiyonu tekrar çalıştırılır
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //sql tablosundaki verileri form üzerindeki textbox ve comboboxlara aktardık.

            textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            textBox6.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            richTextBox1.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            textBox8.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Menu menu = new Menu(); //yeni menü form oluşturduk ve açilmasini sağladık mevcut formuda kapattık
            menu.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e) //güncelleme butonu
        {

            kmt.Connection = bag.baglan();//sql komutuna bağlantı oluşturduk
            kmt.CommandText = "update ilac set barkod_no='"+textBox1.Text+ "',ilacin_adi='"+textBox2.Text+ "',uretici_firma='"+textBox3.Text+ "',kutu_sayisi='"+textBox4.Text+ "',fiyati='"+textBox5.Text+ "',kullanim_amaci='"+textBox6.Text+ "',yan_etkileri='"+richTextBox1.Text+ "', ilac_teslim_alan='"+textBox8.Text+ "' where  id = '" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'";
            kmt.ExecuteNonQuery();
            kmt.Dispose();
            //güncelleme işlemini yapar ardından aşşağıda başarılı mesajı verir ve ilaçlistesini tekrardan doldurur
            MessageBox.Show("İlaç Güncelleme işlemi tamamlandı ! ");
            dtst.Clear();
            ilacListesiDoldur();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            dtst.Clear();
            SqlDataAdapter adtr = new SqlDataAdapter("select * From ilac where barkod_no LIKE '%" + textBox7.Text + "%'", bag.baglan());
            //burada arama işlemi yaptık ancak  normal aramalardan farklı olarak listelerken LIKE Kullandık bu yazdıklarımız eğer tc içerisinde var ise sonuç verecektir
            //bire bir de karşılaştırma yapmayacaktır.
            adtr.Fill(dtst, "ilac");
            dataGridView1.DataMember = "ilac";
            dataGridView1.DataSource = dtst;
            adtr.Dispose();
            //datagridviewi listeledimiz verilerle doldurduk
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // seçilen satırın tamamını seçmesini sağladık ve aşşağıdaki kodlarla kolonlardaki başlıkları düzenledik
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Barkod No";
            dataGridView1.Columns[2].HeaderText = "İlacın Adı";
            dataGridView1.Columns[3].HeaderText = "Üretici Firma";
            dataGridView1.Columns[4].HeaderText = "Kutu Sayısı";
            dataGridView1.Columns[5].HeaderText = "Fiyatı";
            dataGridView1.Columns[6].HeaderText = "Kullanım Amacı";
            dataGridView1.Columns[7].HeaderText = "Yan Etkileri";
            dataGridView1.Columns[8].HeaderText = "İlacı Teslim Alan Personel";
        }

        private void button8_Click(object sender, EventArgs e)
        {

            try
            {


                bag.baglan();
                DialogResult cevap;
                cevap = MessageBox.Show("İlaç Tablosunu da Boşaltılsın İstiyor Musunuz ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes)
                {

                    SqlCommand tablobosalt = new SqlCommand(" delete from ilac", bag.baglan());
                    tablobosalt.ExecuteNonQuery();

                    ///
                }

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                object Missing = Type.Missing;
                Workbook workbook = excel.Workbooks.Add(Missing);
                Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
                int StartCol = 1;
                int StartRow = 1;
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                }
                StartRow++;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {

                        Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                        myRange.Select();


                    }
                }

            }
            catch (Exception hata)
            {

                MessageBox.Show("Hata Aldınız" + hata.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            pageSetupDialog1.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DialogResult pdr = printDialog1.ShowDialog();
            if (pdr == DialogResult.OK)
            {
                printDocument1.Print();
            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs ffff)
        {

            try
            {
                //ÇİZİM BAŞLANGICI
                System.Drawing.Font myFont = new System.Drawing.Font("Calibri", 7); //font oluşturduk
                SolidBrush sbrush = new SolidBrush(Color.Black);//fırça oluşturduk
                Pen myPen = new Pen(Color.Black); //kalem oluşturduk

                ffff.Graphics.DrawString("Düzenlenme Tarihi: " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString(), myFont, sbrush, 50, 25);
                ffff.Graphics.DrawLine(myPen, 50, 45, 770, 45); // Çizgi çizdik... 1. Kalem, 2. X, 3. Y Koordinatı, 4. Uzunluk, 5. BitişX
                //Aşı Vurulan TC , Aşı Vuran TC, Aşı Adı , Etki Süresi , Etkisi, Aşı Vurulma Tarihi
                myFont = new System.Drawing.Font("Calibri", 8, FontStyle.Bold);//Fatura başlığı yazacağımız için fontu kalın yaptık ve puntoyu büyütüp 15 yaptık.
                ffff.Graphics.DrawString("İlaç Tablosu", myFont, sbrush, 375, 65);
                ffff.Graphics.DrawLine(myPen, 50, 95, 770, 95); //çizgi çizdik.

                myFont = new System.Drawing.Font("Calibri", 6, FontStyle.Bold); //Detay başlığını yazacağımız için fontu kalın yapıp puntoyu 10 yaptık.
                ffff.Graphics.DrawString("Barkod No", myFont, sbrush, 50, 110); //Detay başlığı
                ffff.Graphics.DrawString("İlacın Adı", myFont, sbrush, 125, 110); //Detay başlığı
                ffff.Graphics.DrawString("Üretici Firma", myFont, sbrush, 200, 110); // Detay başlığı
                ffff.Graphics.DrawString("Kutu Sayısı", myFont, sbrush, 275, 110); //Detay başlığı
                ffff.Graphics.DrawString("Fiyatı", myFont, sbrush, 350, 110); //Detay başlığı

                ffff.Graphics.DrawString("Kullanım Amacı", myFont, sbrush, 425, 110); //Detay başlığı
                ffff.Graphics.DrawString("Yan Etkileri", myFont, sbrush, 550, 110); //Detay başlığı
                ffff.Graphics.DrawString("İlacı Teslim Alan Personel", myFont, sbrush, 720, 110); //Detay başlığı
                ffff.Graphics.DrawLine(myPen, 25, 125, 770, 125); //Çizgi çizdik.

                int y = 150; //y koordinatının yerini belirledik.(Verilerin yazılmaya başlanacağı yer)

                myFont = new System.Drawing.Font("Calibri", 6); //fontu 10 yaptık.

                int i = 0;//satır sayısı için değişken tanımladık.
                while (i <= dataGridView1.Rows.Count)//döngüyü son satırda sonlandıracağız.
                {
                    ffff.Graphics.DrawString(dataGridView1[1, i].Value.ToString(), myFont, sbrush, 50, y);//1.sütun
                    ffff.Graphics.DrawString(dataGridView1[2, i].Value.ToString(), myFont, sbrush, 125, y);//2.sütun
                    ffff.Graphics.DrawString(dataGridView1[3, i].Value.ToString(), myFont, sbrush, 200, y);//3.sütun
                    ffff.Graphics.DrawString(dataGridView1[4, i].Value.ToString(), myFont, sbrush, 275, y);//4.sütun
                    ffff.Graphics.DrawString(dataGridView1[5, i].Value.ToString() + " TL", myFont, sbrush, 350, y);//5.sütun
                    ffff.Graphics.DrawString(dataGridView1[6, i].Value.ToString(), myFont, sbrush, 425, y);//4.sütun
                    ffff.Graphics.DrawString(dataGridView1[7, i].Value.ToString(), myFont, sbrush, 550, y);//5.sütun
                    ffff.Graphics.DrawString(dataGridView1[8, i].Value.ToString(), myFont, sbrush, 720, y);//5.sütun
                    y += 20; //y koordinatını arttırdık.
                    i += 1; //satır sayısını arttırdık

                    //yeni sayfaya geçme kontrolü
                    if (y > 1000)
                    {
                        ffff.Graphics.DrawString("(Devamı -->)", myFont, sbrush, 700, y + 50);
                        y = 50;
                        break; //burada yazdırma sınırına ulaştığımız için while döngüsünden çıkıyoruz
                               //çıktığımızda while baştan başlıyor i değişkeni değer almaya devam ediyor
                               //yazdırma yeni sayfada başlamış oluyor
                    }
                }
                //çoklu sayfa kontrolü
                if (i < dataGridView1.RowCount - 1)
                {
                    ffff.HasMorePages = true;
                }
                else
                {
                    ffff.HasMorePages = false;
                    i = 0;
                }
                StringFormat myStringFormat = new StringFormat();
                myStringFormat.Alignment = StringAlignment.Far;
            }
            catch
            {
            }
        }
    }
}
